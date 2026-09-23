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
import type { AuditLogsWrite } from '../generated/models/AuditLogsModel'
import { BsGlobe } from "react-icons/bs";
import { CreateOCBCFMSRecordwithNSCService, UploadAttachmenttoApprovalSubmissionsService } from '../generated'
import { getCurrentUser, includesCurrentUser } from './userAccess'
import { ConfirmationDialog } from './ConfirmationDialog'

type Person = { '@odata.type': string; Claims: string; DisplayName: string; Email: string; Picture: string; Department: string; JobTitle: string }
type Option = { id: number; title: string; users?: Person[]; categoryId?: number; categoryTitle?: string }
type StatusMap = Record<string, StatusValue>
type AttachmentTypeMap = Record<string, AttachmentTypeValue>
type ConfirmationAction = 'Send for Approval'
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

export function NewApprovalSubmission({ onBack, onSaved }: { onBack: () => void; onSaved: () => void }) {
  const [form, setForm] = useState<FormState>(emptyForm)
  const [options, setOptions] = useState({ nsc: [] as Option[], functionName: [] as Option[], category: [] as Option[], subcategory: [] as Option[] })
  const [instruction, setInstruction] = useState({ submission: noInstruction, fulfiller: noInstruction, fulfillers: '', fulfillerPeople: [] as Person[] })
  const [statusOptions, setStatusOptions] = useState<string[]>([])
  const [statusMap, setStatusMap] = useState<StatusMap>({})
  const [attachmentTypesOptions, setAttachmentTypesOptions] = useState<string[]>([])
  const [attachmentTypeMap, setAttachmentTypeMap] = useState<AttachmentTypeMap>({})
  const [files, setFiles] = useState<File[]>([])
  const [saving, setSaving] = useState(false)
  const [message, setMessage] = useState('')
  const [validationErrors, setValidationErrors] = useState<RequiredFieldMap>(emptyRequiredMap)
  const [pendingAction, setPendingAction] = useState<ConfirmationAction | null>(null)

  useEffect(() => {
    Promise.all([getCurrentUser(), NSCService.getAll(), FunctionService.getAll(), CategoryService.getAll(), Sub_categoryService.getAll()])
      .then(([currentUser, nsc, functions, categories, subcategories]) => {
        const allNscOptions = nsc.data.map((item) => ({ id: item.ID ?? 0, title: item.Title, users: getPeople(item.Users) }))
        const userNscOptions = allNscOptions.filter((item) => includesCurrentUser(nsc.data.find((nscItem) => nscItem.ID === item.id)?.Users, currentUser.email))
        const fallbackNsc = userNscOptions.length === 0 && allNscOptions.length === 1 ? allNscOptions[0] : null
        setOptions({
          nsc: userNscOptions,
          functionName: functions.data.map((item) => ({ id: item.ID ?? 0, title: item.Title })),
          category: categories.data.map((item) => ({ id: item.ID ?? 0, title: item.Title })),
          subcategory: subcategories.data.map((item) => ({ id: item.ID ?? 0, title: item.Title, categoryId: item.Category?.Id ?? 0, categoryTitle: item.Category?.Value ?? '' })),
        })
        if (fallbackNsc) setForm((current) => ({ ...current, nsc: fallbackNsc }))
      })
      .catch(() => setMessage('Unable to load form values from SharePoint.'))
    ApprovalSubmissionsService.getReferencedEntity('', 'Status')
      .then((result) => {
        const extractedChoices = extractStatusChoices(result.data)
        const values = extractedChoices.map((item) => item.Value)
        const map = extractedChoices.reduce<StatusMap>((accumulator, item) => {
          accumulator[item.Value] = { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: item.Value, Id: item.Id }
          return accumulator
        }, {})
        setStatusOptions(values)
        setStatusMap(map)
        if (values.includes('Draft')) setForm((current) => current.status ? current : { ...current, status: 'Draft' })
      })
      .catch(() => setMessage('Unable to load Status choices from SharePoint.'))
    ApprovalSubmissionsService.getReferencedEntity('', 'AttachmentType')
      .then((result) => {
        const extractedChoices = extractStatusChoices(result.data)
        const values = extractedChoices.map((item) => item.Value)
        const map = extractedChoices.reduce<AttachmentTypeMap>((accumulator, item) => {
          accumulator[item.Value] = { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Value: item.Value, Id: item.Id }
          return accumulator
        }, {})
        setAttachmentTypesOptions(values)
        setAttachmentTypeMap(map)
      })
      .catch(() => setMessage('Unable to load attachment type choices from SharePoint.'))
  }, [])

  useEffect(() => {
    if (!form.nsc || !form.category || !form.subcategory) return
    OCBCFMSApprovalMatrixService.getAll({ filter: `NSC/Title eq '${form.nsc.title.replace(/'/g, "''")}' and Category/Title eq '${form.category.title.replace(/'/g, "''")}' and Subcategory/Title eq '${form.subcategory.title.replace(/'/g, "''")}'` })
      .then((result) => {
        const matrix = result.data[0]
        setInstruction({
          submission: matrix?.SubmissionInstruction || noInstruction,
          fulfiller: matrix?.FulfillerInstruction || noInstruction,
          fulfillers: getPeople(matrix?.Fulfillers).map((person) => person.DisplayName).join(', '),
          fulfillerPeople: getPeople(matrix?.Fulfillers),
        })
      })
      .catch(() => setInstruction({ submission: noInstruction, fulfiller: noInstruction, fulfillers: '', fulfillerPeople: [] }))
  }, [form.nsc, form.category, form.subcategory])

  const reviewers = form.nsc?.users || []
  const setField = <K extends keyof FormState>(key: K, value: FormState[K]) => setForm((current) => ({ ...current, [key]: value }))
  const clearValidation = () => {
    setMessage('')
    setValidationErrors(emptyRequiredMap)
  }
  const requiredFieldMatch = (formState: FormState) => {
    const requiredMissing: RequiredFieldMap = {
      nsc: !formState.nsc,
      functionName: !formState.functionName,
      category: !formState.category,
      subcategory: !formState.subcategory,
      contractingParty: !formState.contractingParty.trim(),
    }
    return requiredMissing
  }
  const selectOption = (key: 'nsc' | 'functionName' | 'category' | 'subcategory', value: string) => {
    const selected = options[key].find((item) => item.title === value) || null
    clearValidation()
    if (key === 'nsc') {
      setField('nsc', selected)
      setField('functionName', null)
      setField('category', null)
      setField('subcategory', null)
      setField('reviewer', null)
      setInstruction({ submission: noInstruction, fulfiller: noInstruction, fulfillers: '', fulfillerPeople: [] })
      return
    }
    if (key === 'functionName') {
      setField('functionName', selected)
      setField('category', null)
      setField('subcategory', null)
      setInstruction({ submission: noInstruction, fulfiller: noInstruction, fulfillers: '', fulfillerPeople: [] })
      return
    }
    if (key === 'category') {
      setField('category', selected)
      setField('subcategory', null)
      setInstruction({ submission: noInstruction, fulfiller: noInstruction, fulfillers: '', fulfillerPeople: [] })
      return
    }
    setField('subcategory', selected)
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

  const uploadAttachments = async (
    itemId: number,
    files: File[]
  ): Promise<void> => {
    for (const file of files) {
      const fileContent = await file.arrayBuffer();
      const bytes = new Uint8Array(fileContent);

      let binary = '';

      const chunkSize = 8192;

      for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode(
          ...bytes.subarray(i, i + chunkSize)
        );
      }

      const base64 = btoa(binary);

      await UploadAttachmenttoApprovalSubmissionsService.Run({
        text: file.name,
        number: itemId,
        file: {
          name: file.name,
          contentBytes: base64
        }
      });
    }
  };

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
      NSC: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: nsc!.id, Value: nsc!.title },
      Function: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: functionName!.id, Value: functionName!.title },
      Category: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: category!.id, Value: category!.title },
      Subcategory: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: subcategory!.id, Value: subcategory!.title },
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
    }
    try {      
      var recordStatus = await CreateOCBCFMSRecordwithNSCService.Run({'text': form.nsc!.title});
      var recordId = recordStatus.data.output;
      const result = await ApprovalSubmissionsService.update(String(recordId), record);
      const createdId = result.data.ID ?? 0
      let generatedTitle = ''
      if (createdId > 0) {
        const now = new Date()
        const datePart = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`
        generatedTitle = `OCBCFMS-${datePart}-${createdId}`
        await ApprovalSubmissionsService.update(String(createdId), { Title: generatedTitle })

        const context = await getContext().catch(() => undefined)
        const actionBy = context?.user?.fullName || context?.user?.userPrincipalName || 'Logged-in user'
        const auditRecord: Omit<AuditLogsWrite, 'ID'> = {
          SubmissionsID: { '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', Id: createdId, Value: generatedTitle },
          Action: actionName,
          ActionBy: actionBy,
          ActionedDateandTime: new Date().toISOString(),
          Comments: '',
        }
        await AuditLogsService.create(auditRecord)
      }
      var itemId = result.data.ID || 0;
      if (files.length > 0) {
        await uploadAttachments(
          itemId,
          files
        );  
      }
      setMessage(requestedStatus === 'Draft' ? 'Approval submission saved as draft.' : 'Approval submission sent for approval.')
      setTimeout(onSaved, 700)
    } catch {
      setMessage('Unable to save the approval submission.');
      onBack;
    } finally {
      setSaving(false)
    }
  }

  const update = (key: 'contractingParty' | 'contractValidity' | 'link' | 'remarks', value: string) => {
    if (key === 'contractingParty') clearValidation()
    setField(key, value)
  }
  const filteredSubcategories = form.category ? options.subcategory.filter((item) => item.categoryId === form.category?.id || item.categoryTitle === form.category?.title) : []
  const visibleInstruction = form.nsc && form.category && form.subcategory ? instruction : { submission: noInstruction, fulfiller: noInstruction, fulfillers: '', fulfillerPeople: [] as Person[] }
  const sendForApproval = () => {
    const requestedStatus = form.reviewer ? 'Pending with Reviewer' : 'Pending with Fulfiller'
    if (!validateSubmission(requestedStatus)) return
    setPendingAction('Send for Approval')
  }

  return <main className="page-content new-submission-page">
    {saving && <div className="page-loader-overlay" role="status" aria-live="polite"><span className="page-loader-spinner" aria-hidden="true" /><span>Saving approval submission...</span></div>}
    <div className="page-heading"><p>New Approval Submission</p><span>Today is {new Date().toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'short', year: 'numeric' })}</span></div>
    <section className="form-section"><h2>Basic Information</h2>
      <div className="form-grid">
        <Field label="NSC " required={true}>{options.nsc.length === 0 ? <input value={form.nsc?.title || ''} readOnly className={validationErrors.nsc ? 'field-error-control' : ''} /> : <Select value={form.nsc?.title || ''} placeholder="Find NSC" options={options.nsc} onChange={(value) => selectOption('nsc', value)} error={validationErrors.nsc} />}</Field>
        <Field label="Function " required={true}><Select value={form.functionName?.title || ''} placeholder="Find function" options={options.functionName} onChange={(value) => selectOption('functionName', value)} error={validationErrors.functionName} /></Field>
        <Field label="Category " required={true}><Select value={form.category?.title || ''} placeholder="Find category" options={options.category} onChange={(value) => selectOption('category', value)} error={validationErrors.category} /></Field>
        <Field label="Sub-category " required={true}><Select value={form.subcategory?.title || ''} placeholder="Find sub-category" options={filteredSubcategories} onChange={(value) => selectOption('subcategory', value)} error={validationErrors.subcategory} /></Field>
        <Field label="Contracting Party Name " required={true}><input className={validationErrors.contractingParty ? 'field-error-control' : ''} value={form.contractingParty} onChange={(e) => update('contractingParty', e.target.value)} placeholder="Enter contracting party name" /></Field>
        <Field label="Contract Validity"><div className="date-control"><input type="date" value={form.contractValidity} onChange={(e) => update('contractValidity', e.target.value)} /><button className="reset-button" onClick={() => update('contractValidity', '')}>↻</button></div></Field>

      </div>
      <div className="form-grid-2 mt-10">
        <Field label="Status" required>
          <select value={form.status || 'Draft'} onChange={(e) => setField('status', e.target.value)} disabled>
            <option value="">Find status</option>{statusOptions.map((status) => <option key={status}>{status}</option>)}
          </select>
        </Field>
        <Field label="Link"><div className="link-control"><input value={form.link} onChange={(e) => update('link', e.target.value)} placeholder="Enter link" /><button className="link-button" onClick={() => form.link && window.open(form.link, '_blank', 'noopener,noreferrer')}><BsGlobe /></button></div></Field>
        <Field label="Remarks" wide><textarea value={form.remarks} onChange={(e) => update('remarks', e.target.value)} placeholder="Enter remarks" /></Field>
      </div>
    </section>
    <section className="form-section"><h2>Attachments</h2>
      <Field label="Attachment Type" required>
        <div className="multi-select">
          <select multiple value={form.attachmentTypes} onChange={(e) => setField('attachmentTypes', Array.from(e.target.selectedOptions, (option) => option.value))}>
            {/* <option value="">Find Attachment Types</option> */}
            {attachmentTypesOptions.map((type) => <option key={type} value={type}>{type}</option>)}
          </select>
        </div>
      </Field>
      <Field label="Attachments">
        <input type="file" multiple onChange={(e) => setFiles(Array.from(e.target.files || []))} /><small>{files.length ? `${files.length} file(s) selected` : 'There is nothing attached.'}</small>
      </Field>
    </section>
    <section className="form-section"><h2>Instructions</h2><Field label="Submission Instruction"><InstructionContent html={visibleInstruction.submission} /></Field><div className="mt-10"><Field label="Fulfiller Instruction"><InstructionContent html={visibleInstruction.fulfiller} /></Field></div></section>
    <section className="form-section"><h2>Approvers</h2><div className="form-grid-2"><Field label="Reviewer"><select value={form.reviewer?.Claims || ''} onChange={(e) => setField('reviewer', reviewers.find((reviewer) => reviewer.Claims === e.target.value) || null)} disabled={!form.nsc}><option value="">Find reviewer</option>{reviewers.map((reviewer) => <option key={reviewer.Claims} value={reviewer.Claims}>{reviewer.DisplayName}</option>)}</select></Field><Field label="Fulfillers (Auto-populated based on the approval matrix)"><input value={visibleInstruction.fulfillers} readOnly /></Field></div></section>
    {message && <div className="form-message">{message}</div>}
    <div className="form-actions">
      <button className="primary-action save-btn" disabled={saving} aria-busy={saving} onClick={() => void save(form.status || 'Draft', 'Save')}>Save</button>
      <button className="primary-action send-for-approval-btn" disabled={saving} aria-busy={saving} onClick={sendForApproval}>Send for Approval</button>
      <button onClick={onBack}>Back</button>
    </div>
    {pendingAction && <ConfirmationDialog action={pendingAction} onConfirm={() => { setPendingAction(null); void save(form.reviewer ? 'Pending with Reviewer' : 'Pending with Fulfiller', 'Send for Approval') }} onCancel={() => setPendingAction(null)} />}
  </main>
}

function Field({ label, required, children, wide = false }: { label: string; required?: boolean; children: ReactNode; wide?: boolean }) {
  return <label className={wide ? 'form-field wide' : 'form-field'}>
    <span className="form-label">{label}{required && <span className="required">*</span>}</span>
    {children}
  </label>
}
function Select({ value, placeholder, options, onChange, error = false }: { value: string; placeholder: string; options: Option[]; onChange: (value: string) => void; error?: boolean }) {
  return <select value={value} className={error ? 'field-error-control' : ''} onChange={(e) => onChange(e.target.value)}>
    <option value="">{placeholder}</option>
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
function InstructionContent({ html }: { html: string }) {
  const sanitizedHtml = html === noInstruction ? '' : sanitizeHtml(html)
  return sanitizedHtml ?
    <div className="instruction-content" dangerouslySetInnerHTML={{ __html: sanitizedHtml }} />
    :
    <div className="instruction-content instruction-empty">{noInstruction}</div>
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
