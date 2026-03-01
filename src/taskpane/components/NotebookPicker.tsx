import React from 'react'
import type { Notebook } from '@/onenote/types'

interface NotebookPickerProps {
  notebooks: Notebook[]
  selectedId: string | null
  onSelect: (notebook: Notebook) => void
  onRefresh: () => void
}

export default function NotebookPicker({
  notebooks,
  selectedId,
  onSelect,
  onRefresh,
}: NotebookPickerProps): React.ReactElement {
  function handleChange(e: React.ChangeEvent<HTMLSelectElement>): void {
    const nb = notebooks.find(n => n.id === e.target.value)
    if (nb) onSelect(nb)
  }

  return (
    <div className="notebook-picker">
      <label className="notebook-picker__label" htmlFor="notebook-select">
        Notebook
      </label>
      <div className="notebook-picker__row">
        <select
          id="notebook-select"
          className="notebook-picker__select"
          value={selectedId ?? ''}
          onChange={handleChange}
        >
          <option value="" disabled>
            — select a notebook —
          </option>
          {notebooks.map(nb => (
            <option key={nb.id} value={nb.id}>
              {nb.displayName}
            </option>
          ))}
        </select>
        <button
          type="button"
          className="notebook-picker__refresh"
          aria-label="Refresh notebook list"
          onClick={onRefresh}
        >
          ↺
        </button>
      </div>
    </div>
  )
}
