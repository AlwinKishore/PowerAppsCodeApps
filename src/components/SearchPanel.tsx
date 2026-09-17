import type { FilterValues } from './types'
import { BsChevronDown, BsChevronUp } from "react-icons/bs";

type Props = {
  open: boolean
  filters: FilterValues
  onToggle: () => void
  onChange: (key: keyof FilterValues, value: string) => void
  onReset: () => void
  onSearch: () => void
  activeTab: string
  onTabChange: (tab: string) => void
  options: Record<'nsc' | 'functionName' | 'category' | 'subcategory' | 'status', string[]>
  subcategoriesByCategory: Record<string, string[]>
}

export function SearchPanel({ open, filters, onToggle, onChange, onReset, onSearch, activeTab, onTabChange, options, subcategoriesByCategory }: Props) {
  const fields: [keyof FilterValues, string, string][] = [['nsc', 'NSC', 'Find NSC'], ['functionName', 'Function', 'Find function'], ['category', 'Category', 'Find category'], ['subcategory', 'Sub-category', 'Find sub-category'], ['contractingParty', 'Contracting Party Name', 'Enter contracting party name'], ['status', 'Status', 'Find status']]
  const subcategoryOptions = filters.category ? subcategoriesByCategory[filters.category] || [] : []
  return <section className={`search-section ${open ? '' : 'collapsed'}`}>
          <div className="search-title">
            <b>Search By</b>
            <button onClick={onToggle}>
              {open ? 'Hide Search' : 'Show Search'}
              <span className="search-action-icons" aria-hidden="true">{open ? <BsChevronUp /> : <BsChevronDown />}</span>
            </button>
          </div>
          <div className="tabs">
              <button className={`tab ${activeTab === 'All' ? 'active' : ''}`} onClick={() => onTabChange('All')}>All</button>
              <button className={`tab ${activeTab !== 'All' ? 'active' : ''}`} onClick={() => onTabChange('My Pending Approvals')}>My Pending Approvals</button>
          </div>
          {open && 
            <><div className="filter-grid">{fields.map(([key, label, placeholder]) => 
                <label key={key}>{label}
                  <div className="input-wrap">
                    {key === 'contractingParty' ? <input value={filters[key]} onChange={(event) => onChange(key, event.target.value)} placeholder={placeholder} /> : <select value={filters[key]} disabled={key === 'subcategory' && !filters.category} onChange={(event) => onChange(key, event.target.value)}><option value="">{placeholder}</option>{(key === 'subcategory' ? subcategoryOptions : options[key]).map((option) => <option key={option} value={option}>{option}</option>)}</select>}
                    {key !== 'contractingParty' && <span className="search-action-icons" aria-hidden="true"><BsChevronDown /></span>}
                  </div>
                </label>
                )}
              </div>
              <div className="filter-actions">
                <button className="reset-button" onClick={onReset}>↻ Reset</button>
                <button className="search-button" onClick={onSearch}>⌕ Search</button>
              </div>
            </>
          }
        </section>
}
