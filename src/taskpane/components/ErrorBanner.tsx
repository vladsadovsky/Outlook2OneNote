import React from 'react'

interface ErrorBannerProps {
  message: string
  onRetry?: () => void
}

export default function ErrorBanner({ message, onRetry }: ErrorBannerProps): React.ReactElement {
  return (
    <div className="error-banner" role="alert">
      <span className="error-banner__message">{message}</span>
      {onRetry && (
        <button type="button" className="error-banner__retry" onClick={onRetry}>
          Retry
        </button>
      )}
    </div>
  )
}
