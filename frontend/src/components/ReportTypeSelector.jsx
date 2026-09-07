import QuarterSelector, { DEALS_QUARTERS, PIPELINE_PERIODS } from './QuarterSelector'

const REPORT_TYPES = [
  { id: 'deals', label: 'Deals', detail: 'Closed / disbursed loans' },
  { id: 'pipeline', label: 'Pipeline', detail: 'Deals still in progress' },
]

const PERIOD_OPTIONS_BY_TYPE = {
  deals: DEALS_QUARTERS,
  pipeline: PIPELINE_PERIODS,
}

export default function ReportTypeSelector({ selected, onSelect, quarter, onQuarterSelect }) {
  return (
    <div className="quarter-selector">
      {REPORT_TYPES.map((r) => (
        <div key={r.id}>
          <button
            className={`quarter-tab ${selected === r.id ? 'active' : ''}`}
            onClick={() => onSelect(r.id)}
          >
            <div>
              <div className="quarter-label">{r.label}</div>
              <div className="quarter-months">{r.detail}</div>
            </div>
          </button>
          {selected === r.id && (
            <QuarterSelector
              selected={quarter}
              onSelect={onQuarterSelect}
              options={PERIOD_OPTIONS_BY_TYPE[r.id]}
            />
          )}
        </div>
      ))}
    </div>
  )
}
