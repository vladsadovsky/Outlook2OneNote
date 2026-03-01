import { describe, it, expect, vi, beforeEach } from 'vitest'
import { createSettingsService, DEFAULT_SETTINGS } from '@/settings/settingsService'
import type { SettingsSchema } from '@/types/settings'

// ---------------------------------------------------------------------------
// Mock Office.context.roamingSettings
// ---------------------------------------------------------------------------
type SaveCallback = (result: { status: string; error?: { message: string } }) => void

let mockStore: Record<string, unknown> = {}

const mockRoamingSettings = {
  get: vi.fn((key: string) => mockStore[key] ?? null),
  set: vi.fn((key: string, value: unknown) => { mockStore[key] = value }),
  saveAsync: vi.fn((callback: SaveCallback) => callback({ status: 'succeeded' })),
}

function setupOffice() {
  ;(globalThis as Record<string, unknown>)['Office'] = {
    context: { roamingSettings: mockRoamingSettings },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
  }
}

beforeEach(() => {
  mockStore = {}
  vi.clearAllMocks()
  mockRoamingSettings.get.mockImplementation((key: string) => mockStore[key] ?? null)
  mockRoamingSettings.set.mockImplementation((key: string, value: unknown) => { mockStore[key] = value })
  mockRoamingSettings.saveAsync.mockImplementation((callback: SaveCallback) =>
    callback({ status: 'succeeded' }),
  )
  setupOffice()
})

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------
describe('createSettingsService', () => {
  it('returns defaults when no stored settings exist', () => {
    const svc = createSettingsService()
    expect(svc.get()).toEqual(DEFAULT_SETTINGS)
  })

  it('merges stored settings over defaults on creation', () => {
    mockStore['o2on_settings'] = { notebookId: 'nb-123', sortOrder: 'desc' } satisfies Partial<SettingsSchema>
    const svc = createSettingsService()
    const settings = svc.get()
    expect(settings.notebookId).toBe('nb-123')
    expect(settings.sortOrder).toBe('desc')
    // Unspecified fields still come from defaults
    expect(settings.preferredLink).toBe(DEFAULT_SETTINGS.preferredLink)
  })

  it('get() returns a copy — mutations do not affect internal state', () => {
    const svc = createSettingsService()
    const copy = svc.get()
    copy.notebookId = 'mutated'
    expect(svc.get().notebookId).toBeNull()
  })

  it('set() merges partial settings into current state', () => {
    const svc = createSettingsService()
    svc.set({ notebookId: 'nb-456', sortOrder: 'desc' })
    const settings = svc.get()
    expect(settings.notebookId).toBe('nb-456')
    expect(settings.sortOrder).toBe('desc')
    expect(settings.preferredLink).toBe(DEFAULT_SETTINGS.preferredLink)
  })

  it('set() does not mutate the object returned by a prior get()', () => {
    const svc = createSettingsService()
    const before = svc.get()
    svc.set({ sortOrder: 'desc' })
    expect(before.sortOrder).toBe('asc')
  })

  it('save() calls roamingSettings.set and saveAsync', async () => {
    const svc = createSettingsService()
    svc.set({ notebookId: 'nb-789' })
    await svc.save()
    expect(mockRoamingSettings.set).toHaveBeenCalledWith('o2on_settings', expect.objectContaining({ notebookId: 'nb-789' }))
    expect(mockRoamingSettings.saveAsync).toHaveBeenCalled()
  })

  it('save() resolves on success', async () => {
    const svc = createSettingsService()
    await expect(svc.save()).resolves.toBeUndefined()
  })

  it('save() rejects when saveAsync returns failure status', async () => {
    mockRoamingSettings.saveAsync.mockImplementation((callback: SaveCallback) =>
      callback({ status: 'failed', error: { message: 'Network error' } }),
    )
    const svc = createSettingsService()
    await expect(svc.save()).rejects.toThrow('Network error')
  })

  it('save() persists the current merged state, not defaults', async () => {
    const svc = createSettingsService()
    svc.set({ notebookId: 'nb-abc', includeAttachments: false })
    await svc.save()
    const saved = mockRoamingSettings.set.mock.calls[0][1] as SettingsSchema
    expect(saved.notebookId).toBe('nb-abc')
    expect(saved.includeAttachments).toBe(false)
  })
})
