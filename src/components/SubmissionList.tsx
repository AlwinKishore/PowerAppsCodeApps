import type { Submission } from './types'

export function SubmissionList({ submissions, onEdit, onView }: { submissions: Submission[]; onEdit?: (submissionId: number) => void; onView?: (submissionId: number) => void }) {
  const value = (text: string) => <span className="submission-value" title={text}>{text}</span>
  const isEditableStatus = (status: string) => { return status.toLowerCase() === 'draft' || status.trim() === 'Pending for Resubmission'} 

  return <section className="list-section">
    <div className="list-scroll">
      {submissions.length === 0 ? 
        <div className="list-message"><i>No data found</i></div> 
        : 
        submissions.map((item) => 
          <article className="submission-card" key={item.id} tabIndex={0} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); isEditableStatus(item.status) ? onEdit?.(item.id) : onView?.(item.id) } }}>
            <div>
              <p><strong className="submission-id" onClick={() => isEditableStatus(item.status) ? onEdit?.(item.id) : onView?.(item.id)} role="button">Submission ID | {value(item.submissionId)}</strong></p>
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
        )
      }
    </div>
  </section>
}
