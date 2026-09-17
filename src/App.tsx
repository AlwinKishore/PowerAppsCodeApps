import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import './App.css'
import { ApprovalSubmissionsService } from './generated/services/ApprovalSubmissionsService'
import type { ApprovalSubmissionsRead } from './generated/models/ApprovalSubmissionsModel'
import { NSCService } from './generated/services/NSCService'
import { FunctionService } from './generated/services/FunctionService'
import { CategoryService } from './generated/services/CategoryService'
import { Sub_categoryService } from './generated/services/Sub_categoryService'
import { Header } from './components/Header'
import { Footer } from './components/Footer'
import { SearchPanel } from './components/SearchPanel'
import { ActionButtons } from './components/ActionButtons'
import { SubmissionList } from './components/SubmissionList'
import { NewApprovalSubmission } from './components/NewApprovalSubmission'
import { EditApprovalSubmission } from './components/EditApprovalSubmission'
import { ViewApprovalSubmission } from './components/ViewApprovalSubmission'
import type { FilterValues, Submission } from './components/types'
import { getCurrentUser, includesCurrentUser } from './components/userAccess'

const mapSubmission = (item: ApprovalSubmissionsRead): Submission => ({
  id: item.ID ?? 0,
  submissionId: item.Title ?? `OCBCFMS-${item.ID ?? ''}`,
  nsc: item.NSC?.Value ?? '-',
  functionName: item.Function?.Value ?? '-',
  category: item.Category?.Value ?? '-',
  subcategory: item.Subcategory?.Value ?? '-',
  contractingParty: item.ContractingPartyName ?? '-',
  contractValidity: item.ContractValidity ?? '-',
  reviewer: item.Reviewer?.DisplayName ?? '-',
  fulfiller: [item.Fulfiller1?.DisplayName, item.Fulfiller2?.DisplayName, item.Fulfiller3?.DisplayName].filter(Boolean).join(', ') || '-',
  submitter: item.Author?.DisplayName ?? '-',
  submittedDate: item.Created ? new Date(item.Created).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '-',
  status: item.Status?.Value ?? '-',
})

const emptyFilters: FilterValues = { nsc: '', functionName: '', category: '', subcategory: '', contractingParty: '', status: '' }

const pendingStatuses = ['Pending with Reviewer', 'Pending with Fulfillers']
const requiredSubmissionFieldsFilter = 'NSC/Title ne null and Function/Title ne null'
const pageSize = 100

const escapeODataValue = (value: string) => value.replace(/'/g, "''")

const buildSharePointFilter = (filters: FilterValues, tab: string) => {
  const clauses = [
    filters.nsc && `NSC/Title eq '${escapeODataValue(filters.nsc)}'`,
    filters.functionName && `Function/Title eq '${escapeODataValue(filters.functionName)}'`,
    filters.category && `Category/Title eq '${escapeODataValue(filters.category)}'`,
    filters.subcategory && `Subcategory/Title eq '${escapeODataValue(filters.subcategory)}'`,
    filters.contractingParty && `ContractingPartyName eq '${escapeODataValue(filters.contractingParty)}'`,
    filters.status && `Status eq '${escapeODataValue(filters.status)}'`,
  ].filter((clause): clause is string => Boolean(clause))

  if (tab !== 'All' && !filters.status) {
    clauses.push(`(Status eq 'Pending with Reviewer' or Status eq 'Pending with Fulfiller' or Status eq 'Pending with Fulfillers')`)
  }

  return [requiredSubmissionFieldsFilter, ...clauses].join(' and ')
}

const matchesCurrentUserOnApprovalSubmission = (item: ApprovalSubmissionsRead, userEmail: string) => {
  if (!userEmail) return true
  const normalizedEmail = userEmail.toLowerCase()
  const people = [item.Author?.Email, item.Reviewer?.Email, item.Fulfiller1?.Email, item.Fulfiller2?.Email, item.Fulfiller3?.Email]
  return people.some((personEmail) => typeof personEmail === 'string' && personEmail.toLowerCase() === normalizedEmail)
}

const belongsToUser = (item: Submission, userName: string) => {
  if (!userName) return false
  const normalizedUserName = userName.toLowerCase()
  return item.reviewer.toLowerCase().includes(normalizedUserName)
    || item.fulfiller.toLowerCase().includes(normalizedUserName)
}

const isPendingStatus = (status: string) => pendingStatuses.some((pendingStatus) =>
  status.toLowerCase().replace('fulfillers', 'fulfiller') === pendingStatus.toLowerCase().replace('fulfillers', 'fulfiller'))

function App() {
  const [submissions, setSubmissions] = useState<Submission[]>([])
  const [activeTab, setActiveTab] = useState('All')
  const [page, setPage] = useState<'list' | 'new' | 'edit' | 'view'>('list')
  const [selectedSubmissionId, setSelectedSubmissionId] = useState<number | null>(null)
  const [searchOpen, setSearchOpen] = useState(false)
  const [filters, setFilters] = useState<FilterValues>(emptyFilters)
  const [connectionMessage, setConnectionMessage] = useState('')
  const [isLoading, setIsLoading] = useState(true)
  const [listPage, setListPage] = useState(1)
  const [totalSubmissions, setTotalSubmissions] = useState<number | null>(null)
  const [hasNextPage, setHasNextPage] = useState(false)
  const [loggedInUser, setLoggedInUser] = useState('')
  const [currentUserEmail, setCurrentUserEmail] = useState('')
  const [filterOptions, setFilterOptions] = useState({ nsc: [], functionName: [], category: [], subcategory: [], status: [] } as Record<'nsc' | 'functionName' | 'category' | 'subcategory' | 'status', string[]>)
  const [subcategoriesByCategory, setSubcategoriesByCategory] = useState<Record<string, string[]>>({})
  const pageTokens = useRef<Record<number, string>>({})

  const loadSubmissions = useCallback(async (nextFilters: FilterValues, nextTab = 'All', nextPage = 1) => {
    setIsLoading(true)
    setConnectionMessage('')
    try {
      const sharePointFilter = buildSharePointFilter(nextFilters, nextTab)
      if (nextPage === 1) pageTokens.current = {}
      const skip = (nextPage - 1) * pageSize
      const result = await ApprovalSubmissionsService.getAll({
        filter: sharePointFilter || undefined,
        orderBy: ['Created desc', 'ID desc'],
        top: pageSize,
        skip,
        count: true,
      })
      const mappedSubmissions = result.data
        .filter((item) => matchesCurrentUserOnApprovalSubmission(item, currentUserEmail))
        .map(mapSubmission)
        .filter((item) => nextTab === 'All' || (isPendingStatus(item.status) && belongsToUser(item, loggedInUser)))
      if (!nextFilters.status) {
        setFilterOptions((current) => ({
          ...current,
          status: nextTab === 'All' ? [...new Set(mappedSubmissions.map((item) => item.status).filter((status) => status !== '-'))].sort() : pendingStatuses,
        }))
      }
      setSubmissions(mappedSubmissions)
      setListPage(nextPage)
      setTotalSubmissions(result.count ?? null)
      setHasNextPage(Boolean(result.skipToken) || result.data.length === pageSize)
      if (result.skipToken) pageTokens.current[nextPage + 1] = result.skipToken
    } catch {
      setConnectionMessage('Unable to load data.')
      setSubmissions([])
    } finally {
      setIsLoading(false)
    }
  }, [currentUserEmail, loggedInUser])

  useEffect(() => {
    let mounted = true
    getCurrentUser()
      .then((currentUser) => {
        if (!mounted) return Promise.reject(new Error('Component unmounted'))
        setLoggedInUser(currentUser.name)
        setCurrentUserEmail(currentUser.email)
        return Promise.all([
          ApprovalSubmissionsService.getAll({ filter: buildSharePointFilter(emptyFilters, 'All'), orderBy: ['Created desc'], top: pageSize, count: true }),
          NSCService.getAll(),
          FunctionService.getAll(),
          CategoryService.getAll(),
          Sub_categoryService.getAll(),
        ]).then((results) => ({ results, userEmail: currentUser.email }))
      })
      .then(({ results, userEmail }) => {
        if (mounted) {
          const [submissionsResult, nscResult, functionResult, categoryResult, subcategoryResult] = results
          const userNscs = nscResult.data.filter((item) => includesCurrentUser(item.Users, userEmail))
          const mappedSubmissions = submissionsResult.data
            .filter((item) => matchesCurrentUserOnApprovalSubmission(item, userEmail))
            .map(mapSubmission)
          setSubmissions(mappedSubmissions)
          setListPage(1)
          setTotalSubmissions(submissionsResult.count ?? null)
          setHasNextPage(Boolean(submissionsResult.skipToken) || submissionsResult.data.length === pageSize)
          pageTokens.current = submissionsResult.skipToken ? { 2: submissionsResult.skipToken } : {}
          setFilterOptions({
            nsc: userNscs.map((item) => item.Title).filter(Boolean).sort(),
            functionName: functionResult.data.map((item) => item.Title).filter(Boolean).sort(),
            category: categoryResult.data.map((item) => item.Title).filter(Boolean).sort(),
            subcategory: subcategoryResult.data.map((item) => item.Title).filter(Boolean).sort(),
            status: [...new Set(mappedSubmissions.map((item) => item.status).filter((status) => status !== '-'))].sort(),
          })
          setSubcategoriesByCategory(subcategoryResult.data.reduce<Record<string, string[]>>((result, item) => {
            const category = item.Category?.Value
            if (!category || !item.Title) return result
            result[category] = [...(result[category] || []), item.Title].sort()
            return result
          }, {}))
        }
      })
      .catch((error) => {
        if (error instanceof Error && error.message === 'Component unmounted') return
        if (mounted) setConnectionMessage('Unable to load filter data.')
      })
      .finally(() => {
        if (mounted) setIsLoading(false)
      })
    return () => { mounted = false }
  }, [loadSubmissions])

  const filteredSubmissions = useMemo(() => submissions.filter((item) => {
    const matches = (value: string, filter: string) => !filter || value.toLowerCase().includes(filter.toLowerCase())
    return matches(item.nsc, filters.nsc) && matches(item.functionName, filters.functionName) && matches(item.category, filters.category)
      && matches(item.subcategory, filters.subcategory) && matches(item.contractingParty, filters.contractingParty) && matches(item.status, filters.status)
      && (activeTab === 'All' || (isPendingStatus(item.status) && belongsToUser(item, loggedInUser)))
  }), [submissions, filters, activeTab, loggedInUser])

  const totalPages = totalSubmissions === null
    ? listPage + (hasNextPage ? 1 : 0)
    : Math.max(1, Math.ceil(totalSubmissions / pageSize))

  const searchOptions = activeTab === 'All' ? filterOptions : { ...filterOptions, status: pendingStatuses }

  const updateFilter = (key: keyof FilterValues, value: string) => {
    const nextFilters = { ...filters, [key]: value, ...(key === 'category' ? { subcategory: '' } : {}) }
    setFilters(nextFilters)
  }
  const resetFilters = () => {
    setFilters(emptyFilters)
    void loadSubmissions(emptyFilters, activeTab)
  }
  const navigateToList = () => {
    void loadSubmissions(filters, activeTab)
    setPage('list')
    window.location.hash = 'list'
    document.getElementById('approval-submission-list')?.scrollIntoView({ behavior: 'smooth' })
  }
  const navigateToNew = () => {
    setPage('new')
    window.location.hash = 'new'
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }
  const navigateToEdit = (submissionId: number) => {
    setSelectedSubmissionId(submissionId)
    setPage('edit')
    window.location.hash = 'edit'
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }
  const navigateToView = (submissionId: number) => {
    setSelectedSubmissionId(submissionId)
    setPage('view')
    window.location.hash = 'view'
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }

  return (
    <div className="app-shell">
      <Header onApprovalSubmissionsClick={navigateToList} />
      {page === 'new' ? <NewApprovalSubmission onBack={navigateToList} onSaved={navigateToList} /> : page === 'edit' && selectedSubmissionId ?
        <EditApprovalSubmission submissionId={selectedSubmissionId} onBack={navigateToList} onSaved={navigateToList} /> : page === 'view' && selectedSubmissionId ?
        <ViewApprovalSubmission submissionId={selectedSubmissionId} onBack={navigateToList} /> : <>
      <main className="page-content">
        <div className="page-heading">
          <p>Approval Submissions</p>
          <span>Today is Wednesday, 09 Sep 2026</span>
        </div>
        <br></br>
        <SearchPanel 
          open={searchOpen} 
          filters={filters} 
          options={searchOptions} 
          subcategoriesByCategory={subcategoriesByCategory}
          activeTab={activeTab} 
          onTabChange={(tab) => { setActiveTab(tab); void loadSubmissions(filters, tab) }} 
          onToggle={() => setSearchOpen((open) => !open)} 
          onChange={updateFilter} 
          onReset={resetFilters} 
          onSearch={() => { setSearchOpen(true); void loadSubmissions(filters, activeTab) }}
        />
        {connectionMessage && <div className="connection-message">{connectionMessage}</div>}
        <div className="results-heading">
          <strong>Search Results</strong>
          <ActionButtons onReset={resetFilters} onNewRequest={navigateToNew} />
        </div>
        {isLoading ? 
            <div id="approval-submission-list" className="list-message">Loading data...</div> 
          : 
            <>
              <div id="approval-submission-list"><SubmissionList submissions={filteredSubmissions} onEdit={navigateToEdit} onView={navigateToView} /></div>
              <div className="pagination" aria-label="Submission list pagination">
                <button type="button" onClick={() => void loadSubmissions(filters, activeTab, listPage - 1)} disabled={listPage === 1 || isLoading}>Previous</button>
                <span>Page {listPage} of {totalPages}</span>
                <button type="button" onClick={() => void loadSubmissions(filters, activeTab, listPage + 1)} disabled={(!hasNextPage && listPage >= totalPages) || isLoading}>Next</button>
              </div>
            </>
        }
      </main>
      </>}
      <Footer />
    </div>
  )
}

export default App
