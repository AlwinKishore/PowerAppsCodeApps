import { useEffect, useState } from 'react'
import { getContext } from '@microsoft/power-apps/app'
import { Office365UsersService } from '../generated/services/Office365UsersService'
type HeaderProps = {
  onApprovalSubmissionsClick: () => void
}

export function Header({ onApprovalSubmissionsClick }: HeaderProps) {
  const [photoUrl, setPhotoUrl] = useState('')
  const [displayName, setDisplayName] = useState('Logged-in user')

  useEffect(() => {
    let mounted = true

    getContext()
      .then(async (context) => {
        if (!mounted) return
        setDisplayName(context.user.fullName || context.user.userPrincipalName || 'Logged-in user')
        if (!context.user.objectId) return

        const result = await Office365UsersService.UserPhoto(context.user.objectId)
        if (mounted && result.data) {
          setPhotoUrl(`data:image/jpeg;base64,${result.data}`)
        }
      })
      .catch((error) => {
        if (mounted) {
          console.warn('Unable to load profile image.', error)
          setPhotoUrl('')
        }
      })

    return () => { mounted = false }
  }, [])

  return <header className="top-bar">
    <div className="brand-bar">
      <strong className="omron-logo">OMRON</strong>
      <div className="brand-actions">
        <button className="approval-nav" onClick={onApprovalSubmissionsClick}>Approval Submissions</button>
        {/* <button className="approval-nav" onClick={onLazyApprovalSubmissionsClick}>Paged Approval Submissions</button> */}
        <span className="user-profile" title={displayName}>
          {photoUrl ? <img src={photoUrl} alt={`${displayName} profile`} onError={() => setPhotoUrl('')} /> : <span className="user-fallback" aria-hidden="true">{displayName.charAt(0).toUpperCase()}</span>}
        </span>
      </div>
    </div>
  </header>
}
