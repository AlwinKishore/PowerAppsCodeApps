export type ConfirmationAction = 'Approve' | 'Reject' | 'Rework' | 'Send for Approval'

type ConfirmationDialogProps = {
  action: ConfirmationAction
  onConfirm: () => void
  onCancel: () => void
}

const actionMessages: Record<ConfirmationAction, string> = {
  Approve: 'Are you sure you want to approve the request?',
  Reject: 'Are you sure you want to reject the request?',
  Rework: 'Are you sure you want to rework the request?',
  'Send for Approval': 'Are you sure you want to send this submission for approval?',
}

export function ConfirmationDialog({ action, onConfirm, onCancel }: ConfirmationDialogProps) {
  return (
    <div className="confirmation-dialog-backdrop" role="presentation" onMouseDown={onCancel}>
      <div
        className="confirmation-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="confirmation-dialog-title"
        aria-describedby="confirmation-dialog-message"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div className="confirmation-dialog-header">
          <h2 id="confirmation-dialog-title">Confirmation</h2>
          <button className="confirmation-dialog-close" type="button" aria-label="Close confirmation dialog" onClick={onCancel}>×</button>
        </div>
        <p id="confirmation-dialog-message">{actionMessages[action]}</p>
        <div className="confirmation-dialog-actions">
          <button className="confirmation-dialog-button confirm" type="button" onClick={onConfirm}>Yes</button>
          <button className="confirmation-dialog-button cancel" type="button" onClick={onCancel}>No</button>
        </div>
      </div>
    </div>
  )
}
