import { useEffect, useRef } from 'react'
import type { Submission } from './types'

type Props = {
  submissions: Submission[]
  currentUser: string
  onEdit?: (submissionId: number) => void
  onView?: (submissionId: number) => void
  onReachEnd?: () => void
  isLoadingMore?: boolean
  hasMore?: boolean
}

export function SubmissionList({ submissions, currentUser, onEdit, onView, onReachEnd, isLoadingMore = false, hasMore = false }: Props) {
  const loadMoreMarker = useRef<HTMLDivElement>(null)
  const value = (text: string) => <span className="submission-value" title={text}>{text}</span>
  const isEditableStatus = (status: string, submitter: string) => {
    const normalizedUser = currentUser.trim().toLowerCase()
    return normalizedUser !== ''
      && submitter.trim().toLowerCase() === normalizedUser
      && ['draft', 'pending for resubmission'].includes(status.trim().toLowerCase())
  }

  useEffect(() => {
    const marker = loadMoreMarker.current
    if (!marker || !onReachEnd || !hasMore || isLoadingMore) return

    const observer = new IntersectionObserver((entries) => {
      if (entries[0]?.isIntersecting) onReachEnd()
    }, { rootMargin: '240px' })
    observer.observe(marker)
    return () => observer.disconnect()
  }, [hasMore, isLoadingMore, onReachEnd])

  return <section className="list-section">
    <div className="list-scroll">
      {submissions.length === 0 ? 
        <div className="list-message"><i>No data found</i></div> 
        : 
        <>{submissions.map((item) => 
          <article className="submission-card" key={item.id} tabIndex={0} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); isEditableStatus(item.status, item.submitter) ? onEdit?.(item.id) : onView?.(item.id) } }}>
            <div>
              <p><strong className="submission-id" onClick={() => isEditableStatus(item.status, item.submitter) ? onEdit?.(item.id) : onView?.(item.id)} role="button">Submission ID | {value(item.submissionId)}</strong></p>
              <p><b>Category</b> | {value(item.category)}</p>
              <p><b>Contract Validity</b> | {value(item.contractValidity)}</p>
              <p><b>Status</b> | {value(item.status)}</p>
            </div>
            <div>
              <p><b>NSC</b> | {value(item.nsc)}</p>
              <p><b>Sub-category</b> | {value(item.subcategory)}</p>
              <p><b>Reviewer</b> | {value(item.reviewer)}</p>
              <p><b>Submitter</b> | {value(item.submitter)}</p>
            </div>
            <div>
              <p><b>Function</b> | {value(item.functionName)}</p>
              <p><b>Contracting Party Name</b> | {value(item.contractingParty)}</p>
              <p><b>Fulfillers</b> | {value(item.fulfiller)}</p>
              <p><b>Submitted Date</b> | {value(item.submittedDate)}</p>
            </div>
          </article>
        )}<div ref={loadMoreMarker} className="lazy-loading-status" aria-live="polite">
          {onReachEnd && (isLoadingMore ? 'Loading more submissions...' : hasMore && 'Scroll to load more')}
        </div></>
      }
    </div>
  </section>
}
