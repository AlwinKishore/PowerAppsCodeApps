import { useCallback, useEffect, useRef, useState } from 'react'
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
import { ApprovalSubmissionsLazyScreen } from './components/ApprovalSubmissionsLazyScreen'
import { NewApprovalSubmission } from './components/NewApprovalSubmission'
import { EditApprovalSubmission } from './components/EditApprovalSubmission'
import { ViewApprovalSubmission } from './components/ViewApprovalSubmission'
import { ApprovalSubmissionDataTable } from './components/ApprovalSubmissionDataTable'
import { DocumentExtractionPage } from './components/DocumentExtractionPage'
import { SupportChatbot } from './components/SupportChatbot'
import type { FilterValues, Submission } from './components/types'
import { getCurrentUser, includesCurrentUser } from './components/userAccess'
import type { IGetAllOptions } from './generated/models/CommonModels'

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

const pendingStatuses = ['Pending with Reviewer', 'Pending with Fulfiller']
const requiredSubmissionFieldsFilter = 'NSC/Title ne null'
const pageSize = 100

const escapeODataValue = (value: string) => value.replace(/'/g, "''")

const buildSharePointFilter = (filters: FilterValues, tab: string, userEmail = '') => {
  const clauses = [
    filters.nsc && `NSC/Title eq '${escapeODataValue(filters.nsc)}'`,
    filters.functionName && `Function/Title eq '${escapeODataValue(filters.functionName)}'`,
    filters.category && `Category/Title eq '${escapeODataValue(filters.category)}'`,
    filters.subcategory && `Subcategory/Title eq '${escapeODataValue(filters.subcategory)}'`,
    filters.contractingParty && `ContractingPartyName eq '${escapeODataValue(filters.contractingParty)}'`,
    filters.status && `Status eq '${escapeODataValue(filters.status) == 'Pending with Fulfillers' ? 'Pending with Fulfiller' : escapeODataValue(filters.status)}'`,
  ].filter((clause): clause is string => Boolean(clause))
  
  if (filters.status === "") {
    if (tab !== 'All' && userEmail ) {
      const escapedEmail = escapeODataValue(userEmail.toLowerCase())
      clauses.push(`(Status eq 'Pending with Reviewer' and Reviewer/EMail eq '${escapedEmail}') or (Status eq 'Pending with Fulfiller' and (Fulfiller1/EMail eq '${escapedEmail}' or Fulfiller2/EMail eq '${escapedEmail}' or Fulfiller3/EMail eq '${escapedEmail}'))`)
    }
  }
  else{
    if (tab !== 'All' && userEmail && filters.status === 'Pending with Reviewer') {
      const escapedEmail = escapeODataValue(userEmail.toLowerCase())
      clauses.push(`(Status eq 'Pending with Reviewer' and Reviewer/EMail eq '${escapedEmail}')`)
    }
    if (tab !== 'All' && userEmail && (filters.status === 'Pending with Fulfillers' || filters.status === 'Pending with Fulfiller')) {
      const escapedEmail = escapeODataValue(userEmail.toLowerCase())
      clauses.push(`(Status eq 'Pending with Fulfiller' and (Fulfiller1/EMail eq '${escapedEmail}' or Fulfiller2/EMail eq '${escapedEmail}' or Fulfiller3/EMail eq '${escapedEmail}'))`)
    }
  }

  if (tab === 'All' && userEmail && clauses.length === 0) {
    const escapedEmail = escapeODataValue(userEmail.toLowerCase())
    clauses.push(`((Status eq 'Draft' or Status eq 'Pending for Resubmission') and Author/EMail eq '${escapedEmail}') or (Status eq 'Pending with Reviewer' or Status eq 'Pending with Fulfiller' or Status eq 'Approved' or Status eq 'Rejected')`)
  }
  // console.log('SharePoint filter:', [requiredSubmissionFieldsFilter, ...clauses].join(' and '));
  return [requiredSubmissionFieldsFilter, ...clauses].join(' and ')
}

const matchesCurrentUserOnApprovalSubmission = (item: ApprovalSubmissionsRead, userEmail: string, tab: string = 'All') => {
  // if (!userEmail) return true
  const normalizedEmail = userEmail.toLowerCase()
  if (item.Status?.Value === 'Pending with Reviewer' && tab !== 'All') {
    return item.Reviewer?.Email?.toLowerCase() === normalizedEmail
  }
  if (item.Status?.Value === 'Pending with Fulfiller' && tab !== 'All') {
    const people = [item.Fulfiller1?.Email, item.Fulfiller2?.Email, item.Fulfiller3?.Email]
    return people.some((personEmail) => typeof personEmail === 'string' && personEmail.toLowerCase() === normalizedEmail)
  }
  if (item.Status?.Value === 'Draft' || item.Status?.Value === 'Pending for Resubmission' && tab == 'All') {
    return item.Author?.Email?.toLowerCase() === normalizedEmail
  }
  return true
}

function App() {
  const [submissions, setSubmissions] = useState<Submission[]>([])
  const [activeTab, setActiveTab] = useState('All')
  const [page, setPage] = useState<'list' | 'lazy-list' | 'table' | 'document-extraction' | 'new' | 'edit' | 'view'>('lazy-list')
  const [selectedSubmissionId, setSelectedSubmissionId] = useState<number | null>(null)
  const [searchOpen, setSearchOpen] = useState(false)
  const [filters, setFilters] = useState<FilterValues>(emptyFilters)
  const [connectionMessage, setConnectionMessage] = useState('')
  const [isLoading, setIsLoading] = useState(true)
  const [isLoadingMore, setIsLoadingMore] = useState(false)
  const [listPage, setListPage] = useState(1)
  const [totalSubmissions, setTotalSubmissions] = useState<number | null>(null)
  const [hasNextPage, setHasNextPage] = useState(false)
  const [loggedInUser, setLoggedInUser] = useState('')
  const [currentUserEmail, setCurrentUserEmail] = useState('')
  const [filterOptions, setFilterOptions] = useState({ nsc: [], functionName: [], category: [], subcategory: [], status: [] } as Record<'nsc' | 'functionName' | 'category' | 'subcategory' | 'status', string[]>)
  const [subcategoriesByCategory, setSubcategoriesByCategory] = useState<Record<string, string[]>>({})
  const pageTokens = useRef<Record<number, string>>({})

  const loadSubmissions = useCallback(async (nextFilters: FilterValues, nextTab = 'All', nextPage = 1, currentPage = 1) => {
    void currentPage
    setIsLoading(true)
    setConnectionMessage('')
    try {
      const sharePointFilter = buildSharePointFilter(nextFilters, nextTab, currentUserEmail)
      // const lastItem = submissions.length > 0 ? submissions[submissions.length - 1] : null;
      let cursorClause = '';
      // if (lastItem) {
      //   cursorClause = (currentPage < nextPage) ? `(ID lt ${lastItem.id})` : `(ID gt ${lastItem.id})`;
      // }
      const updatedSharePointFilter = [sharePointFilter, cursorClause]
        .filter(Boolean)
        .map(clause => `(${clause})`)
        .join(' and ') || undefined;
      if (nextPage === 1) pageTokens.current = {}
      const skip = (nextPage - 1) * pageSize
      const options: IGetAllOptions = {
        filter: updatedSharePointFilter || undefined,
        top: pageSize,
        skip: skip,
        orderBy: ["Created desc", 'ID desc'],
        count: true
      };
      const result = await ApprovalSubmissionsService.getAll(options)
      const mappedSubmissions = result.data
        .filter((item) => matchesCurrentUserOnApprovalSubmission(item, currentUserEmail, nextTab))
        .map(mapSubmission)
        // .filter((item) => nextTab === 'All' || (isPendingStatus(item.status) && belongsToUser(item, loggedInUser)))
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

  const loadMoreSubmissions = useCallback(async (nextFilters: FilterValues, nextTab: string, nextPage: number) => {
    if (isLoadingMore || !hasNextPage) return
    setIsLoadingMore(true)
    setConnectionMessage('')
    try {
      const sharePointFilter = buildSharePointFilter(nextFilters, nextTab, currentUserEmail)
      const lastItem = submissions.length > 0 ? submissions[submissions.length - 1] : null;
      let cursorClause = '';
      if (lastItem) {
        cursorClause = `(ID lt ${lastItem.id})`;
      }

      const updatedSharePointFilter = [sharePointFilter, cursorClause]
        .filter(Boolean)
        .map(clause => `(${clause})`)
        .join(' and ') || undefined;

      const options: IGetAllOptions = {
        filter: updatedSharePointFilter,
        top: pageSize,
        // no `skip` — cursor-based paging replaces it
        orderBy: ["Created desc", "ID desc"],
        count: true
      };
      const result = await ApprovalSubmissionsService.getAll(options)
      const nextSubmissions = result.data
        .filter((item) => matchesCurrentUserOnApprovalSubmission(item, currentUserEmail, nextTab))
        .map(mapSubmission)
      setSubmissions((current) => [...current, ...nextSubmissions])
      setListPage(nextPage)
      setTotalSubmissions(result.count ?? null)
      setHasNextPage(Boolean(result.skipToken) || result.data.length === pageSize)
    } catch {
      setConnectionMessage('Unable to load more data.')
    } finally {
      setIsLoadingMore(false)
    }
  }, [currentUserEmail, hasNextPage, isLoadingMore])

  useEffect(() => {
    let mounted = true
    getCurrentUser()
      .then((currentUser) => {
        if (!mounted) return Promise.reject(new Error('Component unmounted'))
        setLoggedInUser(currentUser.name)
        setCurrentUserEmail(currentUser.email)
        return Promise.all([
          ApprovalSubmissionsService.getAll({ filter: buildSharePointFilter(emptyFilters, 'All', currentUser.email), orderBy: ['Created desc'], top: pageSize, count: true, maxPageSize: pageSize }),
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
            .filter((item) => matchesCurrentUserOnApprovalSubmission(item, userEmail, activeTab))
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
  const onRefresh = () => {
    void loadSubmissions(filters, activeTab)
  }
  const navigateToList = () => {
    void loadSubmissions(filters, activeTab)
    setPage('list')
    window.location.hash = 'list'
    document.getElementById('approval-submission-list')?.scrollIntoView({ behavior: 'smooth' })
  }
  const navigateToLazyList = () => {
    void loadSubmissions(filters, activeTab)
    setPage('lazy-list')
    window.location.hash = 'lazy-list'
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }
  const navigateToTable = () => {
    void loadSubmissions(filters, activeTab)
    setPage('table')
    window.location.hash = 'table'
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }
  const navigateToDocumentExtraction = () => {
    setPage('document-extraction')
    window.location.hash = 'document-extraction'
    window.scrollTo({ top: 0, behavior: 'smooth' })
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
  const onTabChange = (tab: string) => {    
    setSubmissions([]);
    setActiveTab(tab)
    void loadSubmissions(emptyFilters, tab)
    setFilters(emptyFilters)
  }

  return (
    <div className="app-shell">
      <Header onApprovalSubmissionsClick={navigateToLazyList} onApprovalTableClick={navigateToTable} onDocumentExtractionClick={navigateToDocumentExtraction} />
      {page === 'new' ? <NewApprovalSubmission onBack={navigateToList} onSaved={navigateToList} /> : page === 'document-extraction' ?
        <DocumentExtractionPage onBack={navigateToLazyList} /> : page === 'edit' && selectedSubmissionId ?
        <EditApprovalSubmission submissionId={selectedSubmissionId} onBack={navigateToList} onSaved={navigateToList} /> : page === 'view' && selectedSubmissionId ?
        <ViewApprovalSubmission submissionId={selectedSubmissionId} onBack={navigateToList} /> : page === 'lazy-list' ? 
        <ApprovalSubmissionsLazyScreen
          submissions={submissions}
          currentUser={loggedInUser}
          filters={filters}
          activeTab={activeTab}
          searchOpen={searchOpen}
          options={searchOptions}
          subcategoriesByCategory={subcategoriesByCategory}
          connectionMessage={connectionMessage}
          isLoading={isLoading}
          isLoadingMore={isLoadingMore}
          hasMore={hasNextPage}
          onToggleSearch={() => setSearchOpen((open) => !open)}
          onChangeFilter={updateFilter}
          onReset={resetFilters}
          onSearch={() => { setSearchOpen(true); void loadSubmissions(filters, activeTab) }}
          onTabChange={onTabChange}
          onNewRequest={navigateToNew}
          onEdit={navigateToEdit}
          onView={navigateToView}
          onLoadMore={() => { void loadMoreSubmissions(filters, activeTab, listPage + 1) }}
        /> : page === 'table' ? <main className="page-content">
          <div className="page-heading">
            <p>Approval Submissions Table</p>
            <span>Today is {new Date().toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'short', year: 'numeric' })}</span>
          </div>
          {connectionMessage && <div className="connection-message">{connectionMessage}</div>}
          {isLoading ? <div className="list-message">Loading data...</div> : <ApprovalSubmissionDataTable submissions={submissions} currentUser={loggedInUser} onEdit={navigateToEdit} onView={navigateToView} />}
        </main> : <>
      <main className="page-content">
        <div className="page-heading">
          <p>Approval Submissions</p>
          <span>Today is {new Date().toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'short', year: 'numeric' })}</span>
        </div>
        <br></br>
        <SearchPanel 
          open={searchOpen} 
          filters={filters} 
          options={searchOptions} 
          subcategoriesByCategory={subcategoriesByCategory}
          activeTab={activeTab} 
          onTabChange={onTabChange} 
          onToggle={() => setSearchOpen((open) => !open)} 
          onChange={updateFilter} 
          onReset={resetFilters} 
          onSearch={() => { setSearchOpen(true); void loadSubmissions(filters, activeTab) }}
        />
        {connectionMessage && <div className="connection-message">{connectionMessage}</div>}
        <div className="results-heading">
          <strong>Search Results</strong>
          <ActionButtons onReset={onRefresh} onNewRequest={navigateToNew} />
        </div>
        {isLoading ? 
            <div id="approval-submission-list" className="list-message">Loading data...</div> 
          : 
            <>
              <div id="approval-submission-list"><SubmissionList currentUser={loggedInUser} submissions={submissions} onEdit={navigateToEdit} onView={navigateToView} /></div>
              <div className="pagination" aria-label="Submission list pagination">
                <button type="button" onClick={() => void loadSubmissions(filters, activeTab, listPage - 1)} disabled={listPage === 1 || isLoading}>Previous</button>
                <span>Page {listPage} of {totalPages}</span>
                <button type="button" onClick={() => { void loadSubmissions(filters, activeTab, listPage + 1)}} disabled={(!hasNextPage && listPage >= totalPages) || isLoading}>Next</button>
              </div>
            </>
        }
      </main>
      </>}
      <Footer />
      <SupportChatbot />
    </div>
  )
}

export default App
