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
import type { ApprovalSubmissionsWrite, AttachmentTypeValue, StatusValue } from '../generated/models/ApprovalSubmissionsModel'
import type { AuditLogsRead, AuditLogsWrite } from '../generated/models/AuditLogsModel'
import { BsGlobe } from 'react-icons/bs'
import { ApprovalSubmissionAttachmentDeleteService, ApprovalSubmissionAttachmentRetrievalusingItemIDService, UploadAttachmenttoApprovalSubmissionsService } from '../generated'
import { getCurrentUser, includesCurrentUser } from './userAccess'
import { ConfirmationDialog } from './ConfirmationDialog'

type Person = { '@odata.type': string; Claims: string; DisplayName: string; Email: string; Picture: string; Department: string; JobTitle: string }
type Option = { id: number; title: string; users?: Person[]; categoryId?: number; categoryTitle?: string }
type StatusMap = Record<string, StatusValue>
type AttachmentTypeMap = Record<string, AttachmentTypeValue>
type ConfirmationAction = 'Send for Approval'
type SubmissionAttachment = {
  Id: string
  AbsoluteUri: string
  DisplayName: string
  '@odata.type': string
}
type AttachmentFlowResponse = {
  data: any
} | unknown[]
type RequiredFieldKey = 'nsc' | 'functionName' | 'category' | 'subcategory' | 'contractingParty'
type RequiredFieldMap = Record<RequiredFieldKey, boolean>
type FormState = {
  nsc: Option | null; functionName: Option | null; category: Option | null; subcategory: Option | null
  contractingParty: string; contractValidity: string; link: string; remarks: string; attachmentTypes: string[]; reviewer: Person | null
  status: string
}

const emptyRequiredMap: RequiredFieldMap = { nsc: false, functionName: false, category: false, subcategory: false, contractingParty: false }
const emptyForm: FormState = { nsc: null, functionName: null, category: null, subcategory: null, contractingParty: '', contractValidity: '', link: '', remarks: '', attachmentTypes: [], reviewer: null, status: '' }
const noInstruction = 'No instruction have been provided'
var submissionID: number;
//const attachmentFlowUrl = 'https://b4383a0012b1ee988d071d1b024d6c.da.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/29/workflows/c6d4c08b30884bb2987bc90beeb21bc2/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TP97xbvIG4UUFbfhI0OE1zY_4aVEubN9Iq2Du7kIQJE'
const deleteAttachmentFlowUrl = 'https://b4383a0012b1ee988d071d1b024d6c.da.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/19/workflows/6c72e487c018457cb9afb3324b5b98ab/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=81AjAcGKuLYVcwCDfZRhIR78CydfiInUtSYFR8D2IqU'

export function EditApprovalSubmission({ submissionId, onBack, onSaved }: { submissionId: number; onBack: () => void; onSaved: () => void }) {
  const [form, setForm] = useState<FormState>(emptyForm)
  const [options, setOptions] = useState({ nsc: [] as Option[], functionName: [] as Option[], category: [] as Option[], subcategory: [] as Option[] })
  const [attachmentTypesOptions, setAttachmentTypesOptions] = useState<string[]>([])
  const [attachmentTypeMap, setAttachmentTypeMap] = useState<AttachmentTypeMap>({})
  const [attachments, setAttachments] = useState<SubmissionAttachment[]>([])
  const [files, setFiles] = useState<File[]>([])
  const [removedAttachmentIds, setRemovedAttachmentIds] = useState<string[]>([])
  const [instruction, setInstruction] = useState({ submission: noInstruction, fulfiller: noInstruction, fulfillers: '', fulfillerPeople: [] as Person[] })
  const [statusOptions, setStatusOptions] = useState<string[]>([])
  const [statusMap, setStatusMap] = useState<StatusMap>({})
  const [auditLogs, setAuditLogs] = useState<AuditLogsRead[]>([])
  const [message, setMessage] = useState('')
  const [saving, setSaving] = useState(false)
  const [validationErrors, setValidationErrors] = useState<RequiredFieldMap>(emptyRequiredMap)
  const [pendingAction, setPendingAction] = useState<ConfirmationAction | null>(null)

  useEffect(() => {
    Promise.all([
      getCurrentUser(),
      NSCService.getAll(),
      FunctionService.getAll(),
      CategoryService.getAll(),
      Sub_categoryService.getAll(),
      ApprovalSubmissionsService.get(String(submissionId)),
      AuditLogsService.getAll({ filter: `SubmissionsID/Id eq ${submissionId}` }),
      ApprovalSubmissionsService.getReferencedEntity('', 'Status'),
      ApprovalSubmissionsService.getReferencedEntity('', 'AttachmentType'),
    ])
      .then(([currentUser, nsc, functions, categories, subcategories, submission, logs, statusRef, attachments]) => {
        const nscOptions = nsc.data.filter((item) => includesCurrentUser(item.Users, currentUser.email)).map((item) => ({ id: item.ID ?? 0, title: item.Title, users: getPeople(item.Users) }))
        const functionOptions = functions.data.map((item) => ({ id: item.ID ?? 0, title: item.Title }))
        const categoryOptions = categories.data.map((item) => ({ id: item.ID ?? 0, title: item.Title }))
        const subcategoryOptions = subcategories.data.map((item) => ({ id: item.ID ?? 0, title: item.Title, categoryId: item.Category?.Id ?? 0, categoryTitle: item.Category?.Value ?? '' }))
        setOptions({ nsc: nscOptions, functionName: functionOptions, category: categoryOptions, subcategory: subcategoryOptions })

        submissionID = submissionId;
        const item = submission.data
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

        const extractedStatus = extractStatusChoices(statusRef.data)
        const statusValues = extractedStatus.map((item) => item.Value)
        const statusMapValue = extractedStatus.reduce<StatusMap>((accumulator, item) => {
          accumulator[item.Value] = { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: item.Value, Id: item.Id }
          return accumulator
        }, {})
        setStatusOptions(statusValues)
        setStatusMap(statusMapValue)

        const extractedAttachment = extractStatusChoices(attachments.data)
        const attachmentValues = extractedAttachment.map((item) => item.Value)
        const attachmentMap = extractedAttachment.reduce<AttachmentTypeMap>((accumulator, item) => {
          accumulator[item.Value] = { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: item.Value, Id: item.Id }
          return accumulator
        }, {})
        setAttachmentTypesOptions(attachmentValues)
        setAttachmentTypeMap(attachmentMap)

        setAuditLogs(logs.data)

        loadMatrixInstructions(formFromItem, selectedFulfillers)
      })
      .catch(() => setMessage('Unable to load the edit form data.'))
  }, [submissionId])

  useEffect(() => {
    if (!form.nsc || !form.functionName || !form.category || !form.subcategory) return

    OCBCFMSApprovalMatrixService.getAll({
      filter: `NSC/Title eq '${form.nsc.title.replace(/'/g, "''")}' and Category/Title eq '${form.category.title.replace(/'/g, "''")}' and Subcategory/Title eq '${form.subcategory.title.replace(/'/g, "''")}'`,
    })
      .then((result) => {
        const matrix = result.data[0]
        const resolvedFulfillers = getPeople(matrix?.Fulfillers)
        const fallbackFulfillers = instruction.fulfillerPeople.length ? instruction.fulfillerPeople : []
        const fulfillers = resolvedFulfillers.length ? resolvedFulfillers : fallbackFulfillers

        setInstruction({
          submission: matrix?.SubmissionInstruction || noInstruction,
          fulfiller: matrix?.FulfillerInstruction || noInstruction,
          fulfillers: fulfillers.map((person) => person.DisplayName).join(', '),
          fulfillerPeople: fulfillers,
        })
      })
      .catch(() => {
        setInstruction((current) => ({
          submission: current.submission || noInstruction,
          fulfiller: current.fulfiller || noInstruction,
          fulfillers: current.fulfillers,
          fulfillerPeople: current.fulfillerPeople,
        }))
      })
  }, [form.nsc, form.functionName, form.category, form.subcategory])

  const reviewers = form.nsc?.users || []
  const setField = <K extends keyof FormState>(key: K, value: FormState[K]) => {
    setForm((current) => ({ ...current, [key]: value }))
    if (key === 'nsc' || key === 'functionName' || key === 'category' || key === 'subcategory') {
      setValidationErrors(emptyRequiredMap)
    }
  }
  const clearValidation = () => {
    setMessage('')
    setValidationErrors(emptyRequiredMap)
  }
  const requiredFieldMatch = (formState: FormState) => {
    return {
      nsc: !formState.nsc,
      functionName: !formState.functionName,
      category: !formState.category,
      subcategory: !formState.subcategory,
      contractingParty: !formState.contractingParty.trim(),
    }
  }

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

  const uploadAttachments = async (itemId: number, filesToUpload: File[]): Promise<void> => {
    for (const file of filesToUpload) {
      const bytes = new Uint8Array(await file.arrayBuffer())
      let binary = ''
      const chunkSize = 8192
      for (let index = 0; index < bytes.length; index += chunkSize) {
        binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize))
      }

      await UploadAttachmenttoApprovalSubmissionsService.Run({
        text: file.name,
        number: itemId,
        file: { name: file.name, contentBytes: btoa(binary) },
      })
    }
  }

  const update = (key: 'contractingParty' | 'contractValidity' | 'link' | 'remarks', value: string) => {
    if (key === 'contractingParty') clearValidation()
    setField(key, value)
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

  const filteredSubcategories = form.category ? options.subcategory.filter((item) => item.categoryId === form.category?.id || item.categoryTitle === form.category?.title) : []

  const validateSubmission = (requestedStatus: string) => {
    const missingRequired = requiredFieldMatch(form)
    if (Object.values(missingRequired).some(Boolean)) {
      setValidationErrors(missingRequired)
      setMessage('Please select NSC, Function, Category, Sub-category, and Contracting Party.')
      return false
    }
    const nsc = form.nsc
    const functionName = form.functionName
    const category = form.category
    const subcategory = form.subcategory
    if (!nsc || !functionName || !category || !subcategory || !form.contractingParty.trim()) {
      setValidationErrors(requiredFieldMatch(form))
      setMessage('Please select NSC, Function, Category, Sub-category, and Contracting Party.')
      return false
    }

    const desiredStatus = requestedStatus === 'Pending with Reviewer' && !form.reviewer ? 'Pending with Fulfiller' : requestedStatus
    if (desiredStatus === 'Pending with Fulfiller' || desiredStatus === 'Pending with Reviewer') {
      if (!form.reviewer && instruction.fulfillerPeople.length === 0) {
        window.alert('No fulfiller is assigned for the selected NSC, Function, Category, and Sub-category. Please contact the administration.')
        return false
      }
    }
    if (!getStatusObjectFor(desiredStatus)) {
      setMessage('Unable to resolve the selected status in SharePoint.')
      return false
    }
    return true
  }

  const save = async (requestedStatus: string, actionName = 'Save') => {
    if (!validateSubmission(requestedStatus)) return

    const nsc = form.nsc
    const functionName = form.functionName
    const category = form.category
    const subcategory = form.subcategory
    if (!nsc || !functionName || !category || !subcategory || !form.contractingParty.trim()) return
    const desiredStatus = requestedStatus === 'Pending with Reviewer' && !form.reviewer ? 'Pending with Fulfiller' : requestedStatus

    const resolvedStatus = getStatusObjectFor(desiredStatus)
    if (!resolvedStatus) {
      setMessage('Unable to resolve the selected status in SharePoint.')
      return
    }

    const resolveAttachmentTypes = form.attachmentTypes
      .map((type) => getAttachmentTypeObjectFor(type))
      .filter((v): v is AttachmentTypeValue => Boolean(v))

    setSaving(true)
    setMessage('')

    const record: Omit<ApprovalSubmissionsWrite, 'ID'> = {
      NSC: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: nsc.id, Value: nsc.title },
      Function: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: functionName.id, Value: functionName.title },
      Category: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: category.id, Value: category.title },
      Subcategory: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: subcategory.id, Value: subcategory.title },
      ContractingPartyName: form.contractingParty.trim(),
      ContractValidity: form.contractValidity || undefined,
      Link: form.link,
      Remarks: form.remarks,
      Reviewer: form.reviewer || undefined,
      Status: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: resolvedStatus.Value, Id: resolvedStatus.Id },
      AttachmentType: resolveAttachmentTypes,
      Fulfiller1: instruction.fulfillerPeople[0],
      Fulfiller2: instruction.fulfillerPeople[1],
      Fulfiller3: instruction.fulfillerPeople[2],
      ReviewerComments: '',
      FulfillerComments: '',
      ReviewerActionedDate: '',
      FulfillerActionedDate: '',
      FulfillerActionedBy: undefined,
    }

    try {
      await ApprovalSubmissionsService.update(String(submissionId), record)
      await deleteRemovedAttachments(removedAttachmentIds)
      await uploadAttachments(submissionId, files)

      const context = await getContext().catch(() => undefined)
      const actionBy = context?.user?.fullName || context?.user?.userPrincipalName || 'Logged-in user'
      const auditRecord: Omit<AuditLogsWrite, 'ID'> = {
        SubmissionsID: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: submissionId, Value: '' },
        Action: actionName,
        ActionBy: actionBy,
        ActionedDateandTime: new Date().toISOString(),
        Comments: '',
      }
      await AuditLogsService.create(auditRecord)

      setMessage(actionName === 'Send for Approval' ? 'Approval submission sent for approval.' : 'Approval submission updated.')
      setTimeout(onSaved, 700)
    } catch {
      setMessage('Unable to update the approval submission.')
    } finally {
      setSaving(false)
    }
  }

  const sendForApproval = () => {
    const requestedStatus = form.reviewer ? 'Pending with Reviewer' : 'Pending with Fulfiller'
    if (!validateSubmission(requestedStatus)) return
    setPendingAction('Send for Approval')
  }

  const removeAttachment = (attachment: SubmissionAttachment) => {
    setAttachments((current) => current.filter((item) => item.Id !== attachment.Id))
    setRemovedAttachmentIds((current) => current.includes(attachment.Id) ? current : [...current, attachment.Id])
  }

  return <main className="page-content new-submission-page">
    {saving && <div className="page-loader-overlay" role="status" aria-live="polite"><span className="page-loader-spinner" aria-hidden="true" /><span>Updating approval submission...</span></div>}
    <div className="page-heading"><p>Edit Approval Submission</p><span>Today is Wednesday, 09 Sep 2026</span></div>
    <section className="form-section"><h2>Basic Information</h2>
      <div className="form-grid">
        <Field label="NSC" required={true}><input value={form.nsc?.title || ''} readOnly className={validationErrors.nsc ? 'field-error-control' : ''} /></Field>
        <Field label="Function" required={true}><Select value={form.functionName?.title || ''} placeholder="Find function" options={options.functionName} onChange={(value) => setField('functionName', options.functionName.find((item) => item.title === value) || null)} error={validationErrors.functionName} /></Field>
        <Field label="Category" required={true}><Select value={form.category?.title || ''} placeholder="Find category" options={options.category} onChange={(value) => setField('category', options.category.find((item) => item.title === value) || null)} error={validationErrors.category} /></Field>
        <Field label="Sub-category" required={true}><Select value={form.subcategory?.title || ''} placeholder="Find sub-category" options={filteredSubcategories} onChange={(value) => setField('subcategory', filteredSubcategories.find((item) => item.title === value) || null)} error={validationErrors.subcategory} /></Field>
        <Field label="Contracting Party Name" required={true}><input className={validationErrors.contractingParty ? 'field-error-control' : ''} value={form.contractingParty} onChange={(e) => update('contractingParty', e.target.value)} placeholder="Enter contracting party name" /></Field>
        <Field label="Contract Validity"><div className="date-control"><input type="date" value={form.contractValidity} onChange={(e) => update('contractValidity', e.target.value)} /><button className="reset-button" onClick={() => update('contractValidity', '')}>↻</button></div></Field>
      </div>
      <div className="form-grid-2 mt-10">
        <Field label="Status" required>
          <select value={form.status || 'Draft'} onChange={(e) => setField('status', e.target.value)} disabled>
            <option value="">Find status</option>{statusOptions.map((status) => <option key={status}>{status}</option>)}
          </select>
        </Field>
        <Field label="Link"><div className='link-control'><input value={form.link} onChange={(e) => update('link', e.target.value)} placeholder='Enter link' /><button className='link-button' onClick={() => form.link && window.open(form.link, '_blank', 'noopener,noreferrer')}><BsGlobe /></button></div></Field>
        <Field label="Remarks" wide><textarea value={form.remarks} onChange={(e) => update('remarks', e.target.value)} placeholder='Enter remarks' /></Field>
      </div>
    </section>
    <section className='form-section'><h2>Attachments</h2>
      <Field label='Attachment Type' required>
        <div className='multi-select'>
          <select multiple value={form.attachmentTypes} onChange={(e) => setField('attachmentTypes', Array.from(e.target.selectedOptions, (option) => option.value))}>
            {attachmentTypesOptions.map((type) => <option key={type} value={type}>{type}</option>)}
          </select>
        </div>
      </Field>
      <div className="mt-10">
      <Field label='Attachments'><input type="file" multiple onChange={(event) => setFiles(Array.from(event.target.files || []))} /><small>{files.length ? `${files.length} file(s) selected` : 'There is nothing attached.'}</small></Field>      
      <Field label=''>
        {attachments.length ? <div className="submission-attachments">{attachments.map((attachment) => <div key={attachment.Id} className="submission-attachment-item"><a href={attachment.AbsoluteUri} download={attachment.DisplayName} target="_blank" rel="noopener noreferrer" className="submission-attachment-link">{attachment.DisplayName}</a><button type="button" className="submission-attachment-remove" aria-label={`Remove ${attachment.DisplayName}`} onClick={() => removeAttachment(attachment)}>×</button></div>)}</div> : <small>No attachments</small>}
      </Field>
      </div>
    </section>
    <section className='form-section'><h2>Instructions</h2><Field label='Submission Instruction'><InstructionContent html={instruction.submission} /></Field><div className='mt-10'><Field label='Fulfiller Instruction'><InstructionContent html={instruction.fulfiller} /></Field></div></section>
    <section className='form-section'><h2>Approvers</h2><div className='form-grid-2'><Field label='Reviewer'><select value={form.reviewer?.Claims || ''} onChange={(e) => setField('reviewer', reviewers.find((reviewer) => reviewer.Claims === e.target.value) || null)} disabled={!form.nsc}><option value=''>Find reviewer</option>{reviewers.map((reviewer) => <option key={reviewer.Claims} value={reviewer.Claims}>{reviewer.DisplayName}</option>)}</select></Field><Field label='Fulfillers (Auto-populated based on the approval matrix)'><input value={instruction.fulfillers} readOnly /></Field></div></section>
    {message && <div className='form-message'>{message}</div>}
    <div className='form-actions'>
      <button className='primary-action save-btn' disabled={saving} aria-busy={saving} onClick={() => void save(form.status || 'Draft', 'Save')}>Save</button>
      <button className='primary-action send-for-approval-btn' disabled={saving} aria-busy={saving} onClick={sendForApproval}>Send for Approval</button>
      <button onClick={onBack}>Back</button>
    </div>
    {pendingAction && <ConfirmationDialog action={pendingAction} onConfirm={() => { setPendingAction(null); void save(form.reviewer ? 'Pending with Reviewer' : 'Pending with Fulfiller', 'Send for Approval') }} onCancel={() => setPendingAction(null)} />}
    <section className='form-section audit-section'><h2>Audit Logs</h2>
      <div className='audit-log-table'>
        <table>
          <thead>
            <tr><th>Action</th><th>Action By</th><th>Actioned Date and Time</th><th>Comments</th></tr>
          </thead>
          <tbody>
            {auditLogs.length === 0 ? <tr><td colSpan={4}>No audit logs found.</td></tr> : auditLogs.map((log) => <tr key={log.ID}><td>{log.Action}</td><td>{log.ActionBy}</td><td>{log.ActionedDateandTime ? new Date(log.ActionedDateandTime).toLocaleString("en-GB", {
              day: "2-digit",
              month: "short",
              year: "numeric",
              hour: "2-digit",
              minute: "2-digit",
              hour12: true,
            }) : "-"}</td><td>{log.Comments}</td></tr>)}
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

function Select({ value, placeholder, options, onChange, error = false }: { value: string; placeholder: string; options: Option[]; onChange: (value: string) => void; error?: boolean }) {
  return <select value={value} className={error ? 'field-error-control' : ''} onChange={(e) => onChange(e.target.value)}>
    <option value=''>{placeholder}</option>
    {options.map((option) => <option key={option.id}>{option.title}</option>)}
  </select>
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
  for (const key of ['value', 'Value', 'choices', 'Choices', 'items', 'Items']) {
    if (key in container) return container[key]
  }
  return value
}

function extractAttachments(value: { Id?: string; AbsoluteUri?: string; DisplayName?: string; '@odata.type'?: string }[] | undefined): SubmissionAttachment[] {
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

async function deleteRemovedAttachments(attachmentIds: string[]): Promise<void> {
  if (attachmentIds.length === 0) return
  if (!deleteAttachmentFlowUrl) {
    throw new Error('VITE_DELETE_ATTACHMENT_FLOW_URL is not configured.')
  }
  await Promise.all(attachmentIds.map((attachmentId) => deleteAttachmentFromFlow(submissionID, attachmentId)))
}

async function deleteAttachmentFromFlow(submissionID: number, attachmentId: string): Promise<void> {
  // const response = await fetch(deleteAttachmentFlowUrl as string, {
  //   method: 'POST',
  //   headers: {
  //     'Content-Type': 'application/json',
  //     Accept: 'application/json',
  //   },
  //   body: JSON.stringify({ itemid: submissionID, attachmentid: attachmentId }),
  // })
  // if (!response.ok) {
  //   throw new Error(`Attachment delete flow request failed with status ${response.status}`)
  // }
  await ApprovalSubmissionAttachmentDeleteService.Run({ number: submissionID, text: String(attachmentId) })
}
async function loadAttachmentsFromFlow(
  itemId: number,
  onLoaded: (attachments: SubmissionAttachment[]) => void
): Promise<void> {
  // const response = await fetch(attachmentFlowUrl, {
  //   method: 'POST',
  //   headers: {
  //     'Content-Type': 'application/json',
  //     'Accept': 'application/json',
  //   },
  //   body: JSON.stringify({
  //     itemid: itemId,
  //   }),
  // });
  const response = await ApprovalSubmissionAttachmentRetrievalusingItemIDService.Run({ text : String(itemId) })

  // if (!response.ok) {
  //   throw new Error(
  //     `Attachment flow request failed with status ${response.status}`
  //   );
  // }

  const result = await response as AttachmentFlowResponse;

  onLoaded(extractFlowAttachments(result))
}

function extractFlowAttachments(value: AttachmentFlowResponse): SubmissionAttachment[] {
  const candidates = Array.isArray(value)
    ? value
    : value && typeof value === 'object'
      ? value?.data?.attachments
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
