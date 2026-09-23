import DataTable from 'react-data-table-component'
import type { TableColumn } from 'react-data-table-component'
import type { Submission } from './types'

type ApprovalSubmissionDataTableProps = {
  submissions: Submission[]
  currentUser: string
  onEdit?: (submissionId: number) => void
  onView?: (submissionId: number) => void
}

const isEditableStatus = (status: string, submitter: string, currentUser: string) => {
  const normalizedUser = currentUser.trim().toLowerCase()
  return normalizedUser !== ''
    && submitter.trim().toLowerCase() === normalizedUser
    && ['draft', 'pending for resubmission'].includes(status.trim().toLowerCase())
}

const dateSort = (rowA: Submission, rowB: Submission) => {
  const dateA = Date.parse(rowA.submittedDate)
  const dateB = Date.parse(rowB.submittedDate)
  return (Number.isNaN(dateA) ? 0 : dateA) - (Number.isNaN(dateB) ? 0 : dateB)
}

export function ApprovalSubmissionDataTable({ submissions, currentUser, onEdit, onView }: ApprovalSubmissionDataTableProps) {
  const columns: TableColumn<Submission>[] = [
    {
      name: 'Submission ID',
      selector: (row) => row.submissionId,
      sortable: true,
      cell: (row) => <button className="data-table-link" type="button" onClick={() => isEditableStatus(row.status, row.submitter, currentUser) ? onEdit?.(row.id) : onView?.(row.id)}>{row.submissionId}</button>,
      grow: 1.2,
    },
    { name: 'NSC', selector: (row) => row.nsc, sortable: true, grow: 1.1 },
    { name: 'Function', selector: (row) => row.functionName, sortable: true, grow: 1.1 },
    { name: 'Category', selector: (row) => row.category, sortable: true, grow: 1.1 },
    { name: 'Sub-category', selector: (row) => row.subcategory, sortable: true, grow: 1.2 },
    { name: 'Contracting Party', selector: (row) => row.contractingParty, sortable: true, grow: 1.3 },
    { name: 'Reviewer', selector: (row) => row.reviewer, sortable: true, grow: 1.2 },
    { name: 'Fulfiller', selector: (row) => row.fulfiller, sortable: true, grow: 1.4 },
    { name: 'Submitted Date', selector: (row) => row.submittedDate, sortable: true, sortFunction: dateSort, grow: 1.1 },
    {
      name: 'Status',
      selector: (row) => row.status,
      sortable: true,
      cell: (row) => <span className={`data-table-status status-${row.status.toLowerCase().replace(/\s+/g, '-')}`}>{row.status}</span>,
      grow: 1.2,
    },
    {
      name: 'Action',
      button: true,
      cell: (row) => <button className="data-table-action" type="button" onClick={() => isEditableStatus(row.status, row.submitter, currentUser) ? onEdit?.(row.id) : onView?.(row.id)}>{isEditableStatus(row.status, row.submitter, currentUser) ? 'Edit' : 'View'}</button>,
      ignoreRowClick: true,
      allowOverflow: true,
    },
  ]

  return <section className="approval-data-table-section" aria-label="Approval submissions data table">
    <DataTable
      columns={columns}
      data={submissions}
      pagination
      paginationPerPage={10}
      paginationRowsPerPageOptions={[10, 25, 50]}
      highlightOnHover
      responsive
      persistTableHead
      noDataComponent={<div className="data-table-empty">No approval submissions found.</div>}
      customStyles={{
        headCells: { style: { backgroundColor: '#edf7f8', color: '#27778a', fontWeight: 700, fontSize: '13px' } },
        cells: { style: { color: '#222', fontSize: '13px', paddingTop: '10px', paddingBottom: '10px' } },
        rows: { style: { minHeight: '48px' } },
      }}
    />
  </section>
}
