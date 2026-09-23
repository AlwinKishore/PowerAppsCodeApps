import { ActionButtons } from './ActionButtons'
import { SearchPanel } from './SearchPanel'
import { SubmissionList } from './SubmissionList'
import type { FilterValues, Submission } from './types'

type FilterOptions = Record<'nsc' | 'functionName' | 'category' | 'subcategory' | 'status', string[]>

type Props = {
  submissions: Submission[]
  currentUser: string
  filters: FilterValues
  activeTab: string
  searchOpen: boolean
  options: FilterOptions
  subcategoriesByCategory: Record<string, string[]>
  connectionMessage: string
  isLoading: boolean
  isLoadingMore: boolean
  hasMore: boolean
  onToggleSearch: () => void
  onChangeFilter: (key: keyof FilterValues, value: string) => void
  onReset: () => void
  onSearch: () => void
  onTabChange: (tab: string) => void
  onNewRequest: () => void
  onEdit: (submissionId: number) => void
  onView: (submissionId: number) => void
  onLoadMore: () => void
}

export function ApprovalSubmissionsLazyScreen({
  submissions,
  currentUser,
  filters,
  activeTab,
  searchOpen,
  options,
  subcategoriesByCategory,
  connectionMessage,
  isLoading,
  isLoadingMore,
  hasMore,
  onToggleSearch,
  onChangeFilter,
  onReset,
  onSearch,
  onTabChange,
  onNewRequest,
  onEdit,
  onView,
  onLoadMore,
}: Props) {
  return <main className="page-content lazy-list-screen">
    <div className="page-heading">
      <p>Approval Submissions</p>
      <span>Today is {new Date().toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'short', year: 'numeric' })}</span>
    </div>
    <br />
    <SearchPanel
      open={searchOpen}
      filters={filters}
      options={options}
      subcategoriesByCategory={subcategoriesByCategory}
      activeTab={activeTab}
      onTabChange={onTabChange}
      onToggle={onToggleSearch}
      onChange={onChangeFilter}
      onReset={onReset}
      onSearch={onSearch}
    />
    {connectionMessage && <div className="connection-message">{connectionMessage}</div>}
    <div className="results-heading">
      <strong>Search Results</strong>
      <ActionButtons onReset={onReset} onNewRequest={onNewRequest} />
    </div>
    {isLoading ? <div className="list-message">Loading data...</div> : <>
      <SubmissionList currentUser={currentUser} submissions={submissions} onEdit={onEdit} onView={onView} onReachEnd={onLoadMore} isLoadingMore={isLoadingMore} hasMore={hasMore} />
    </>}
  </main>
}
