import { describe, it, expect } from 'vitest'
import { formatDateForSection, formatDateTimeForPageTitle } from '@/utils/dateFormatter'

describe('formatDateForSection', () => {
  it('formats an ISO date string as YYYY-MM-DD', () => {
    expect(formatDateForSection('2024-03-15T10:30:00Z')).toMatch(/^\d{4}-\d{2}-\d{2}$/)
  })

  it('uses local date components', () => {
    // Create a date that is midnight UTC (so different local days in +/- offset zones)
    // We just verify the format is correct, not the exact value
    const result = formatDateForSection('2024-06-01T00:00:00.000Z')
    expect(result).toMatch(/^2024-0[56]-\d{2}$/)
  })

  it('pads single-digit months and days', () => {
    // 2024-01-05 UTC — local date may differ, but padding is tested
    const result = formatDateForSection('2024-01-05T12:00:00.000Z')
    expect(result).toMatch(/^\d{4}-\d{2}-\d{2}$/)
    const parts = result.split('-')
    expect(parts[1].length).toBe(2)
    expect(parts[2].length).toBe(2)
  })
})

describe('formatDateTimeForPageTitle', () => {
  it('returns YYYY-MM-DD HH:MM format', () => {
    const result = formatDateTimeForPageTitle('2024-07-20T14:05:00.000Z')
    expect(result).toMatch(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/)
  })

  it('pads single-digit hours and minutes', () => {
    const result = formatDateTimeForPageTitle('2024-07-20T08:03:00.000Z')
    const [, time] = result.split(' ')
    const [h, m] = time.split(':')
    expect(h.length).toBe(2)
    expect(m.length).toBe(2)
  })
})
