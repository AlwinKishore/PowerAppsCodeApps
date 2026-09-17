import { useEffect, useState } from 'react'
import type { ReactNode } from 'react'
import { getContext } from '@microsoft/power-apps/app'
import { NSCService } from '../generated/services/NSCService'
import { FunctionService } from '../generated/services/FunctionService'
import { CategoryService } from '../generated/services/CategoryService'
import { Sub_categoryService } from '../generated/services/Sub_categoryService'
import { ApprovalSubmissionsService } from '../generated/services/ApprovalSubmissionsService'
import { AuditLogsService } from '../generated/services/AuditLogsService'
import { OCBCFMSApprovalMatrixService } from '../generated/services/OCBCFMSApprovalMatrixService'
import type { ApprovalSubmissionsRead, ApprovalSubmissionsWrite, AttachmentTypeValue, StatusValue } from '../generated/models/ApprovalSubmissionsModel'
import type { AuditLogsRead, AuditLogsWrite } from '../generated/models/AuditLogsModel'
import { BsGlobe } from 'react-icons/bs'
import { getCurrentUser, includesCurrentUser } from './userAccess'

type Person = { '@odata.type': string; Claims: string; DisplayName: string; Email: string; Picture: string; Department: string; JobTitle: string }
type Option = { id: number; title: string; users?: Person[]; categoryId?: number; categoryTitle?: string }
type StatusMap = Record<string, StatusValue>
type AttachmentTypeMap = Record<string, AttachmentTypeValue>
type SubmissionAttachment = {
  Id: string
  AbsoluteUri: string
  DisplayName: string
  '@odata.type': string
}
type AttachmentFlowResponse = { attachments?: unknown } | unknown[]

const attachmentFlowUrl = 'https://b4383a0012b1ee988d071d1b024d6c.da.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/29/workflows/c6d4c08b30884bb2987bc90beeb21bc2/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TP97xbvIG4UUFbfhI0OE1zY_4aVEubN9Iq2Du7kIQJE'

type FormState = {
  nsc: Option | null; functionName: Option | null; category: Option | null; subcategory: Option | null
  contractingParty: string; contractValidity: string; link: string; remarks: string; attachmentTypes: string[]; reviewer: Person | null
  status: string
}

const noInstruction = 'No instruction have been provided'
const emptyForm: FormState = { nsc: null, functionName: null, category: null, subcategory: null, contractingParty: '', contractValidity: '', link: '', remarks: '', attachmentTypes: [], reviewer: null, status: '' }

export function ViewApprovalSubmission({ submissionId, onBack }: { submissionId: number; onBack: () => void }) {
  const [form, setForm] = useState<FormState>(emptyForm)
  const [attachmentTypeMap, setAttachmentTypeMap] = useState<AttachmentTypeMap>({})
  const [attachments, setAttachments] = useState<SubmissionAttachment[]>([])
  const [instruction, setInstruction] = useState({ submission: noInstruction, fulfiller: noInstruction, fulfillers: '', fulfillerPeople: [] as Person[] })
  const [statusMap, setStatusMap] = useState<StatusMap>({})
  const [auditLogs, setAuditLogs] = useState<AuditLogsRead[]>([])
  const [message, setMessage] = useState('')
  const [saving, setSaving] = useState(false)
  const [reviewerComments, setReviewerComments] = useState('')
  const [fulfillerComments, setFulfillerComments] = useState('')
  const [reviewerActionedDate, setReviewerActionedDate] = useState('')
  const [fulfillerActionedDate, setFulfillerActionedDate] = useState('')
  const [fulfillerActionedBy, setFulfillerActionedBy] = useState<Person | null>(null)
  const [contextUser, setContextUser] = useState<{name:string, email:string, principal:string, claims:string}>({name:'', email:'', principal:'', claims:''})
  const [currentContext, setCurrentContext] = useState<{ user?: { fullName?: string; email?: string; userPrincipalName?: string } } | null>(null)
  const [commentMode, setCommentMode] = useState<{reviewerEditable:boolean, fulfillerEditable:boolean}>({reviewerEditable:false, fulfillerEditable:false})

  useEffect(() => {
    Promise.all([
      getCurrentUser(),
      NSCService.getAll(),
      FunctionService.getAll(),
      CategoryService.getAll(),
      Sub_categoryService.getAll(),
      ApprovalSubmissionsService.getAll({ filter: `ID eq ${submissionId}` }),
      AuditLogsService.getAll({ filter: `SubmissionsID/Id eq ${submissionId}` }),
      ApprovalSubmissionsService.getReferencedEntity('', 'Status'),
      ApprovalSubmissionsService.getReferencedEntity('', 'AttachmentType'),
      getContext().catch(() => undefined)
    ])
      .then(([currentUser, nsc, functions, categories, subcategories, submission, logs, statusRef, attachments, context]) => {
        const nscOptions = nsc.data.filter((item) => includesCurrentUser(item.Users, currentUser.email)).map((item) => ({ id: item.ID ?? 0, title: item.Title, users: getPeople(item.Users) }))
        const functionOptions = functions.data.map((item) => ({ id: item.ID ?? 0, title: item.Title }))
        const categoryOptions = categories.data.map((item) => ({ id: item.ID ?? 0, title: item.Title }))
        const subcategoryOptions = subcategories.data.map((item) => ({ id: item.ID ?? 0, title: item.Title, categoryId: item.Category?.Id ?? 0, categoryTitle: item.Category?.Value ?? '' }))

        const item = submission.data[0] as ApprovalSubmissionsRead | undefined
        if (!item) {
          setMessage('Unable to find the selected approval submission.')
          return
        }
        setAttachments(extractAttachments(item['{Attachments}']))
        void loadAttachmentsFromFlow(submissionId, setAttachments)
          .catch(() => setMessage('Unable to load attachments.'))
        const selectedFulfillers = [item.Fulfiller1, item.Fulfiller2, item.Fulfiller3].filter((person): person is Person => Boolean(person))
        const selectedFulfillersText = selectedFulfillers.map((person) => person.DisplayName).filter(Boolean).join(', ')
        setInstruction({
          submission: noInstruction,
          fulfiller: noInstruction,
          fulfillers: selectedFulfillersText,
          fulfillerPeople: selectedFulfillers,
        })
        const statusValue = item.Status?.Value ?? ''
        const formFromItem: FormState = {
          nsc: nscOptions.find((opt) => opt.title === item.NSC?.Value) ?? null,
          functionName: functionOptions.find((opt) => opt.title === item.Function?.Value) ?? null,
          category: categoryOptions.find((opt) => opt.title === item.Category?.Value) ?? null,
          subcategory: subcategoryOptions.find((opt) => opt.title === item.Subcategory?.Value) ?? null,
          contractingParty: item.ContractingPartyName ?? '',
          contractValidity: item.ContractValidity ?? '',
          link: item.Link ?? '',
          remarks: item.Remarks ?? '',
          attachmentTypes: item.AttachmentType ? item.AttachmentType.map((choice) => choice.Value) : [],
          reviewer: item.Reviewer ?? null,
          status: statusValue,
        }
        setForm(formFromItem)
        setReviewerComments(item.ReviewerComments ?? '')
        setFulfillerComments(item.FulfillerComments ?? '')
        setReviewerActionedDate(item.ReviewerActionedDate ?? '')
        setFulfillerActionedDate(item.FulfillerActionedDate ?? '')
        setFulfillerActionedBy(item.FulfillerActionedBy ?? null)

        const extractedStatus = extractStatusChoices(statusRef.data)
        const statusMapValue = extractedStatus.reduce<StatusMap>((accumulator, item) => {
          accumulator[item.Value] = { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: item.Value, Id: item.Id }
          return accumulator
        }, {})
        setStatusMap(statusMapValue)

        const extractedAttachment = extractStatusChoices(attachments.data)
        const attachmentMap = extractedAttachment.reduce<AttachmentTypeMap>((accumulator, item) => {
          accumulator[item.Value] = { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: item.Value, Id: item.Id }
          return accumulator
        }, {})
        setAttachmentTypeMap(attachmentMap)

        setAuditLogs(logs.data)

        const userFullName = context?.user?.fullName || ''
        const userPrincipal = context?.user?.userPrincipalName || ''
        const claims = context?.user?.userPrincipalName || ''
        setCurrentContext(context ?? null)
        setContextUser({
          name: userFullName,
          email: userPrincipal,
          principal: userPrincipal,
          claims,
        })

        const reviewerEditable = statusValue.toLowerCase() === 'pending with reviewer' && !!item.Reviewer && personMatchesCurrentUser(item.Reviewer, context)
        const fulfillerEditable = statusValue.toLowerCase() === 'pending with fulfiller' && selectedFulfillers.some((person) => personMatchesCurrentUser(person, context))
        setCommentMode({ reviewerEditable, fulfillerEditable })

        loadMatrixInstructions(formFromItem, selectedFulfillers)
      })
      .catch(() => setMessage('Unable to load the view form data.'))
  }, [submissionId])

  const getStatusObjectFor = (choiceValue: string): StatusValue | undefined => {
    const key = Object.keys(statusMap).find((statusName) => statusName.toLowerCase() === choiceValue.toLowerCase())
    if (!key) return undefined
    return statusMap[key]
  }

  const getAttachmentTypeObjectFor = (choiceValue: string): AttachmentTypeValue | undefined => {
    const key = Object.keys(attachmentTypeMap).find((attachmentTypeName) => attachmentTypeName.toLowerCase() === choiceValue.toLowerCase())
    if (!key) return undefined
    return attachmentTypeMap[key]
  }

  const loadMatrixInstructions = (formState: FormState, fallbackFulfillers: Person[] = []) => {
    if (!formState.nsc || !formState.functionName || !formState.category || !formState.subcategory) return

    OCBCFMSApprovalMatrixService.getAll({
      filter: `NSC/Title eq '${formState.nsc.title.replace(/'/g, "''")}' and Category/Title eq '${formState.category.title.replace(/'/g, "''")}' and Subcategory/Title eq '${formState.subcategory.title.replace(/'/g, "''")}'`,
    })
      .then((result) => {
        const matrix = result.data[0]
        const resolvedFulfillers = getPeople(matrix?.Fulfillers)
        const selectedFulfillers = resolvedFulfillers.length ? resolvedFulfillers : fallbackFulfillers

        setInstruction({
          submission: matrix?.SubmissionInstruction || noInstruction,
          fulfiller: matrix?.FulfillerInstruction || noInstruction,
          fulfillers: selectedFulfillers.map((person) => person.DisplayName).join(', '),
          fulfillerPeople: selectedFulfillers,
        })
      })
      .catch(() => {
        setInstruction({
          submission: noInstruction,
          fulfiller: noInstruction,
          fulfillers: fallbackFulfillers.map((person) => person.DisplayName).join(', '),
          fulfillerPeople: fallbackFulfillers,
        })
      })
  }

  const updateSubmission = async (_requestedStatus: string, actionName: 'Approve' | 'Reject' | 'Rework') => {
    if (saving) return
    const commentRequired = commentMode.reviewerEditable && actionName !== 'Approve' && !reviewerComments.trim()
    if (commentRequired) {
      setMessage('Reviewer comments are required before this action.')
      return
    }
    const fulfillerRequired = commentMode.fulfillerEditable && actionName !== 'Approve' && !fulfillerComments.trim()
    if (fulfillerRequired) {
      setMessage('Fulfiller comments are required before this action.')
      return
    }

    const statusValue = form.status
    const desiredStatus = getStatusObjectFor(statusValue)
    if (!desiredStatus) {
      setMessage('Unable to resolve the selected status in SharePoint.')
      return
    }

    const baseRecord: Omit<ApprovalSubmissionsWrite, 'ID'> = {
      NSC: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: form.nsc?.id ?? 0, Value: form.nsc?.title ?? '' },
      Function: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: form.functionName?.id ?? 0, Value: form.functionName?.title ?? '' },
      Category: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: form.category?.id ?? 0, Value: form.category?.title ?? '' },
      Subcategory: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: form.subcategory?.id ?? 0, Value: form.subcategory?.title ?? '' },
      ContractingPartyName: form.contractingParty.trim(),
      ContractValidity: form.contractValidity || undefined,
      Link: form.link,
      Remarks: form.remarks,
      Reviewer: form.reviewer || undefined,
      Status: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: resolvedStatusValueForAction(actionName, statusValue, form.reviewer, instruction.fulfillerPeople.length).Value, Id: resolvedStatusValueForAction(actionName, statusValue, form.reviewer, instruction.fulfillerPeople.length).Id },
      AttachmentType: form.attachmentTypes.map((type) => getAttachmentTypeObjectFor(type)).filter((v): v is AttachmentTypeValue => Boolean(v)),
      Fulfiller1: instruction.fulfillerPeople[0],
      Fulfiller2: instruction.fulfillerPeople[1],
      Fulfiller3: instruction.fulfillerPeople[2],
    }

    const targetStatus = resolvedStatusValueForAction(actionName, statusValue, form.reviewer, instruction.fulfillerPeople.length)
    const next = { ...baseRecord, Status: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: targetStatus.Value, Id: targetStatus.Id } } as Omit<ApprovalSubmissionsWrite, 'ID'>

    if (statusValue.toLowerCase() === 'pending with reviewer' && actionName === 'Approve') {
      next.ReviewerComments = reviewerComments.trim()
      next.ReviewerActionedDate = new Date().toISOString()
      if (instruction.fulfillerPeople.length > 0) {
        next.Status = getStatusObjectFor('Pending with Fulfiller')
      } else {
        next.Status = getStatusObjectFor('Approved')
      }
    } else if (statusValue.toLowerCase() === 'pending with reviewer' && actionName === 'Reject') {
      if (!reviewerComments.trim()) {
        setMessage('Reviewer comments are mandatory while rejecting.')
        return
      }
      next.ReviewerComments = reviewerComments.trim()
      next.ReviewerActionedDate = new Date().toISOString()
      next.Status = getStatusObjectFor('Reject')
    } else if (statusValue.toLowerCase() === 'pending with reviewer' && actionName === 'Rework') {
      if (!reviewerComments.trim()) {
        setMessage('Reviewer comments are mandatory while reworking.')
        return
      }
      next.ReviewerComments = reviewerComments.trim()
      next.ReviewerActionedDate = new Date().toISOString()
      next.Status = getStatusObjectFor('Pending for Resubmission')
    } else if (statusValue.toLowerCase() === 'pending with fulfiller' && actionName === 'Approve') {
      next.FulfillerComments = fulfillerComments.trim()
      next.FulfillerActionedDate = new Date().toISOString()
      next.FulfillerActionedBy = contextUserToPerson(contextUser)
      next.Status = getStatusObjectFor('Approved')
    } else if (statusValue.toLowerCase() === 'pending with fulfiller' && actionName === 'Reject') {
      if (!fulfillerComments.trim()) {
        setMessage('Fulfiller comments are mandatory while rejecting.')
        return
      }
      next.FulfillerComments = fulfillerComments.trim()
      next.FulfillerActionedDate = new Date().toISOString()
      next.FulfillerActionedBy = contextUserToPerson(contextUser)
      next.Status = getStatusObjectFor('Rejected')
    } else if (statusValue.toLowerCase() === 'pending with fulfiller' && actionName === 'Rework') {
      if (!fulfillerComments.trim()) {
        setMessage('Fulfiller comments are mandatory while reworking.')
        return
      }
      next.FulfillerComments = fulfillerComments.trim()
      next.FulfillerActionedDate = new Date().toISOString()
      next.FulfillerActionedBy = contextUserToPerson(contextUser)
      next.Status = getStatusObjectFor('Pending for Resubmission')
    }

    try {
      setSaving(true)
      setMessage('')
      await ApprovalSubmissionsService.update(String(submissionId), next)

      const context = await getContext().catch(() => undefined)
      const actionBy = context?.user?.fullName || context?.user?.userPrincipalName || 'Logged-in user'
      const action = actionName === 'Approve' && form.status === "Pending with Reviewer" ? 'Approved (Reviewer)' : actionName === 'Approve' && form.status === "Pending with Fulfiller" ? 'Approved (Fulfiller)' : actionName === 'Reject' && form.status === "Pending with Reviewer" ? 'Rejected (Reviewer)' : actionName === 'Reject' && form.status === "Pending with Fulfiller" ? 'Rejected (Fulfiller)' :  actionName === 'Rework' && form.status === "Pending with Reviewer" ? 'Rework (Reviewer)' : 'Rework (Fulfiller)'
      const auditRecord: Omit<AuditLogsWrite, 'ID'> = {
        SubmissionsID: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: submissionId, Value: '' },
        Action: action,
        ActionBy: actionBy,
        ActionedDateandTime: new Date().toISOString(),
        Comments: form.status == "Pending with Reviewer" ? reviewerComments.trim() ? reviewerComments.trim() : '' : fulfillerComments.trim() ? fulfillerComments.trim() : '',
      }
      await AuditLogsService.create(auditRecord)

      setMessage(`${actionName} action completed.`)
      setTimeout(() => onBack(), 500)
    } catch {
      setMessage('Unable to apply the approval action.')
    } finally {
      setSaving(false)
    }
  }

  const resolvedStatusValueForAction = (actionName: 'Approve' | 'Reject' | 'Rework', currentStatus: string, _reviewer: Person | null, fulfillerCount: number): StatusValue => {
    if (actionName === 'Approve' && currentStatus.toLowerCase() === 'pending with reviewer') {
      return fulfillerCount > 0 ? getStatusObjectFor('Pending with Fulfiller') || { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: 'Pending with Fulfiller', Id: 0 } : getStatusObjectFor('Approved') || { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: 'Approved', Id: 0 }
    }
    if (actionName === 'Approve' && currentStatus.toLowerCase() === 'pending with fulfiller') {
      return getStatusObjectFor('Approved') || { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: 'Approved', Id: 0 }
    }
    if (actionName === 'Reject') {
      return getStatusObjectFor('Reject') || { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: 'Reject', Id: 0 }
    }
    if (actionName === 'Rework') {
      return getStatusObjectFor('Pending for Resubmission') || { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: 'Pending for Resubmission', Id: 0 }
    }
    return { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: currentStatus, Id: 0 }
  }

  const isTerminalApprovalStatus = (status: string) => {
    const normalized = (status || '').toLowerCase().trim()
    return normalized === 'approved' || normalized === 'rejected' || normalized === 'reject'
  }

  const reviewerActionAllowed = !isTerminalApprovalStatus(form.status)
    && form.status.toLowerCase().trim() === 'pending with reviewer'
    && !!form.reviewer
    && !!currentContext
    && personMatchesCurrentUser(form.reviewer, currentContext)

  const fulfillerActionAllowed = !isTerminalApprovalStatus(form.status)
    && form.status.toLowerCase() === 'pending with fulfiller'
    && instruction.fulfillerPeople.some((person) => !!currentContext && personMatchesCurrentUser(person, currentContext))

  const showActionButtons = reviewerActionAllowed || fulfillerActionAllowed

  return <main className="page-content new-submission-page">
    {saving && <div className="page-loader-overlay" role="status" aria-live="polite"><span className="page-loader-spinner" aria-hidden="true" /><span>Processing approval action...</span></div>}
    <div className="page-heading"><p>View Approval Submission</p><span>Today is {new Date().toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'short', year: 'numeric' })}</span></div>
    <section className="form-section"><h2>Basic Information</h2>
      <div className="form-grid">
        <Field label="Submission ID"><input value={form.nsc?.title ? `OCBCFMS-${submissionId}` : `OCBCFMS-${submissionId}`} readOnly /></Field>
        <Field label="Submitter"><input value={contextUser.name || ''} readOnly /></Field>
        <Field label="Submitted Date"><input value={new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })} readOnly /></Field>
        <Field label="NSC"><input value={form.nsc?.title || ''} readOnly /></Field>
        <Field label="Function"><input value={form.functionName?.title || ''} readOnly /></Field>
        <Field label="Category"><input value={form.category?.title || ''} readOnly /></Field>
        <Field label="Sub-category"><input value={form.subcategory?.title || ''} readOnly /></Field>
        <Field label="Contracting Party Name"><input value={form.contractingParty} readOnly /></Field>
        <Field label="Contract Validity"><input type="date" value={form.contractValidity} readOnly /></Field>
        <Field label="Status"><input value={form.status} readOnly /></Field>
        <Field label="Link"><div className='link-control'><input value={form.link} readOnly /><button className='link-button' onClick={() => form.link && window.open(form.link, '_blank', 'noopener,noreferrer')}><BsGlobe /></button></div></Field>
        <Field label="Remarks" wide><textarea value={form.remarks} readOnly placeholder='Enter remarks' /></Field>
      </div>
    </section>

    <section className='form-section'><h2>Attachments</h2>
      <Field label='Attachment Type'>
        <input value={form.attachmentTypes.length ? form.attachmentTypes.join(', ') : ''} readOnly />
      </Field>
      <div className="mt-10">        
      <Field label='Attachments'>
        {attachments.length ? <div className="submission-attachments">{attachments.map((attachment) => <a key={attachment.Id || attachment.AbsoluteUri} href={attachment.AbsoluteUri} download={attachment.DisplayName} target="_blank" rel="noopener noreferrer" className="submission-attachment-link">{attachment.DisplayName}</a>)}</div> : <small>No attachments</small>}
      </Field>
      </div>
    </section>

    <section className='form-section'><h2>Instructions</h2><Field label='Submission Instruction'><InstructionContent html={instruction.submission} /></Field><div className='mt-10'><Field label='Fulfiller Instruction'><InstructionContent html={instruction.fulfiller} /></Field></div></section>

    <section className='form-section'><h2>Approvers</h2>
      <div className='approver-grid'>
        <div className='approver-row'>
          <span className='approver-label'>Reviewer</span>
          <span className='approver-value'>{form.reviewer?.DisplayName || '—'}</span>
          <span className='approver-label'>Reviewer Actioned Date</span>
          <span className='approver-value'>{reviewerActionedDate ? new Date(reviewerActionedDate).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true }) : '—'}</span>
        </div>
        <div className='approver-row comments-row'>
          <span className='approver-label'>Reviewer Comments</span>
          <textarea className={commentMode.reviewerEditable ? 'comment-box' : 'comment-box readonly'} value={reviewerComments} readOnly={!commentMode.reviewerEditable} onChange={(e) => setReviewerComments(e.target.value)} placeholder='Enter reviewer comments' />
        </div>
        <div className='approver-row'>
          <span className='approver-label'>Fulfillers</span>
          <span className='approver-value'>{instruction.fulfillers || '—'}</span>
          <span className='approver-label'>Fulfiller Actioned By</span>
          <span className='approver-value'>{fulfillerActionedBy?.DisplayName || '—'}</span>
          <span className='approver-label'>Fulfiller Actioned Date</span>
          <span className='approver-value'>{fulfillerActionedDate ? new Date(fulfillerActionedDate).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true }) : '—'}</span>
        </div>
        <div className='approver-row comments-row'>
          <span className='approver-label'>Fulfiller Comments</span>
          <textarea className={commentMode.fulfillerEditable ? 'comment-box' : 'comment-box readonly'} value={fulfillerComments} readOnly={!commentMode.fulfillerEditable} onChange={(e) => setFulfillerComments(e.target.value)} placeholder='Enter fulfiller comments' />
        </div>
      </div>
    </section>

    {message && <div className='form-message'>{message}</div>}
    <section className='action-section'>
      <div className='action-section-buttons'>
        {isTerminalApprovalStatus(form.status) || !showActionButtons ? (
          <button className='action-button back-button' onClick={onBack}>Back</button>
        ) : (
          <>
            <button className='action-button approve-button' disabled={saving} aria-busy={saving} onClick={() => void updateSubmission(form.status, 'Approve')}>Approve</button>
            <button className='action-button reject-button' disabled={saving} aria-busy={saving} onClick={() => void updateSubmission(form.status, 'Reject')}>Reject</button>
            <button className='action-button rework-button' disabled={saving} aria-busy={saving} onClick={() => void updateSubmission(form.status, 'Rework')}>Rework</button>
            <button className='action-button back-button' onClick={onBack}>Back</button>
          </>
        )}
      </div>
    </section>

    <section className='form-section audit-section'><h2>Audit Logs</h2>
      <div className='audit-log-table'>
        <table>
          <thead>
            <tr><th>Action</th><th>Action By</th><th>Actioned Date and Time</th><th>Comments</th></tr>
          </thead>
          <tbody>
            {auditLogs.length === 0 ? <tr><td colSpan={4}>No audit logs found.</td></tr> : auditLogs.map((log) => <tr key={log.ID}><td>{log.Action}</td><td>{log.ActionBy}</td><td>{log.ActionedDateandTime ? new Date(log.ActionedDateandTime).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true }) : '-'}</td><td>{log.Comments}</td></tr>)}
          </tbody>
        </table>
      </div>
    </section>
  </main>
}

function Field({ label, required, children, wide = false }: { label: string; required?: boolean; children: ReactNode; wide?: boolean }) {
  return <label className={wide ? 'form-field wide' : 'form-field'}>
    <span className='form-label'>{label}{required && <span className='required'>*</span>}</span>
    {children}
  </label>
}

function InstructionContent({ html }: { html: string }) {
  const sanitizedHtml = html === noInstruction ? '' : sanitizeHtml(html)
  return sanitizedHtml ?
    <div className='instruction-content' dangerouslySetInnerHTML={{ __html: sanitizedHtml }} />
    :
    <div className='instruction-content instruction-empty'>{noInstruction}</div>
}

function sanitizeHtml(html: string) {
  const documentFragment = new DOMParser().parseFromString(html, 'text/html')
  documentFragment.querySelectorAll('script, style, iframe, object, embed').forEach((node) => node.remove())
  documentFragment.querySelectorAll('*').forEach((element) => {
    Array.from(element.attributes).forEach((attribute) => {
      if (attribute.name.toLowerCase().startsWith('on') || attribute.name.toLowerCase() === 'srcdoc') element.removeAttribute(attribute.name)
    })
  })
  return documentFragment.body.innerHTML
}

function getPeople(value: unknown): Person[] {
  const values = Array.isArray(value) ? value : value ? [value] : []
  return values.filter((person): person is Person => typeof person === 'object' && person !== null && 'Claims' in person && 'DisplayName' in person)
}

function extractStatusChoices(value: unknown): Array<{ Value: string; Id: number }> {
  const candidates = Array.isArray(value) ? value : unwrapChoiceCollection(value)
  if (!Array.isArray(candidates)) return []
  return candidates.map((item) => {
    if (item && typeof item === 'object') {
      const choice = item as { Value?: unknown; value?: unknown; Label?: unknown; label?: unknown; Id?: unknown; id?: unknown }
      const text = [choice.Value, choice.value, choice.Label, choice.label]
        .find((candidate): candidate is string => typeof candidate === 'string')
      const id = typeof choice.Id === 'number' ? choice.Id : typeof choice.id === 'number' ? choice.id : 0
      return text ? { Value: text, Id: id } : undefined
    }
    return undefined
  }).filter((item): item is { Value: string; Id: number } => Boolean(item) && typeof item?.Value === 'string')
}

function unwrapChoiceCollection(value: unknown): unknown {
  if (!value || typeof value !== 'object') return value
  const container = value as Record<string, unknown>
  const choices = container.value || container.data || container.results || container.items || []
  return Array.isArray(choices) ? choices : []
}

function extractAttachments(value: ApprovalSubmissionsRead['{Attachments}']): SubmissionAttachment[] {
  return (value || []).reduce<SubmissionAttachment[]>((result, attachment) => {
    if (!attachment.AbsoluteUri || !attachment.DisplayName) return result
    result.push({
      Id: attachment.Id || attachment.AbsoluteUri,
      AbsoluteUri: attachment.AbsoluteUri,
      DisplayName: attachment.DisplayName,
      '@odata.type': attachment['@odata.type'] || '#Microsoft.Azure.Connectors.SharePoint.SPListItemAttachment',
    })
    return result
  }, [])
}

async function loadAttachmentsFromFlow(
  itemId: number,
  onLoaded: (attachments: SubmissionAttachment[]) => void
): Promise<void> {
  const response = await fetch(attachmentFlowUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
    },
    body: JSON.stringify({
      itemid: itemId,
    }),
  });

  if (!response.ok) {
    throw new Error(
      `Attachment flow request failed with status ${response.status}`
    );
  }

  const result = await response.json() as AttachmentFlowResponse;

  onLoaded(extractFlowAttachments(result))
}

function extractFlowAttachments(value: AttachmentFlowResponse): SubmissionAttachment[] {
  const candidates = Array.isArray(value)
    ? value
    : value && typeof value === 'object'
      ? value.attachments
      : []
  if (!Array.isArray(candidates)) return []
  return candidates.reduce<SubmissionAttachment[]>((result, candidate) => {
    if (!candidate || typeof candidate !== 'object') return result
    const attachment = candidate as {
      Id?: unknown
      AbsoluteUri?: unknown
      DisplayName?: unknown
      '@odata.type'?: unknown
    }
    if (typeof attachment.Id !== 'string' || typeof attachment.AbsoluteUri !== 'string' || typeof attachment.DisplayName !== 'string') return result
    result.push({
      Id: attachment.Id,
      AbsoluteUri: attachment.AbsoluteUri,
      DisplayName: attachment.DisplayName,
      '@odata.type': typeof attachment['@odata.type'] === 'string'
        ? attachment['@odata.type']
        : '#Microsoft.Azure.Connectors.SharePoint.SPListItemAttachment',
    })
    return result
  }, [])
}

function personMatchesCurrentUser(person: Person | null | undefined, context: any) {
  if (!person || !context?.user) return false
  const userPrincipal = context.user.userPrincipalName || ''
  const personEmail = (person.Email || '').toLowerCase()

  return personEmail.includes(userPrincipal)
}

function contextUserToPerson(current: {name:string, email:string, principal:string, claims:string}): Person {
  return {
    '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference',
    Claims: current.claims || current.principal,
    DisplayName: current.name || current.principal,
    Email: current.email || current.principal,
    Picture: '',
    Department: '',
    JobTitle: '',
  }
}
