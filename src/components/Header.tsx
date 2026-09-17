import { useEffect, useState } from 'react'
import { getContext } from '@microsoft/power-apps/app'

type HeaderProps = {
  onApprovalSubmissionsClick: () => void
}

export function Header({ onApprovalSubmissionsClick }: HeaderProps) {
  const [photoUrl, setPhotoUrl] = useState('')
  const [displayName, setDisplayName] = useState('Logged-in user')

  useEffect(() => {
    let mounted = true

    getContext()
      .then((context) => {
        if (!mounted) return
        setDisplayName(context.user.fullName || context.user.userPrincipalName || 'Logged-in user')
        if (context.user.objectId) {
          setPhotoUrl(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(context.user.objectId)}/photo/$value`)
        }
      })
      .catch(() => {
        // The local preview host does not provide Power Apps user context.
      })

    return () => { mounted = false }
  }, [])

  return <header className="top-bar">
    <div className="brand-bar">
      <strong className="omron-logo">OMRON</strong>
      <div className="brand-actions">
        <button className="approval-nav" onClick={onApprovalSubmissionsClick}>Approval Submissions</button>
        <span className="user-profile" title={displayName}>
          {photoUrl ? <img src={photoUrl} alt={`${displayName} profile`} onError={() => setPhotoUrl('')} /> : <span className="user-fallback" aria-hidden="true">{displayName.charAt(0).toUpperCase()}</span>}
        </span>
      </div>
    </div>
  </header>
}
