import React from 'react'

interface ProgressBarProps {
  message: string
}

export default function ProgressBar({ message }: ProgressBarProps): React.ReactElement {
  return (
    <div className="progress-bar" role="status" aria-live="polite">
      <div className="progress-bar__spinner" aria-hidden="true" />
      <span className="progress-bar__message">{message}</span>
    </div>
  )
}
