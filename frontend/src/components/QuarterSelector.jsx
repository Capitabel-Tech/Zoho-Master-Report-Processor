export const DEALS_QUARTERS = [
  { id: 'Q1', label: 'Q1', months: 'Apr / May / Jun' },
  { id: 'Q2', label: 'Q2', months: 'Jul / Aug / Sep' },
  { id: 'Q3', label: 'Q3', months: 'Oct / Nov / Dec' },
  { id: 'Q4', label: 'Q4', months: 'Jan / Feb / Mar' },
]

export const PIPELINE_PERIODS = [
  { id: 'FULL', label: 'Pipeline Deals for Q2', months: '' },
  { id: 'LT40', label: 'Pipeline Deals for Q2 - <40%', months: '' },
]

export default function QuarterSelector({ selected, onSelect, options = DEALS_QUARTERS }) {
  return (
    <select
      className="quarter-dropdown"
      value={selected}
      onChange={(e) => onSelect(e.target.value)}
      onClick={(e) => e.stopPropagation()}
    >
      {options.map((q) => (
        <option key={q.id} value={q.id}>
          {q.months ? `${q.label} — ${q.months}` : q.label}
        </option>
      ))}
    </select>
  )
}
