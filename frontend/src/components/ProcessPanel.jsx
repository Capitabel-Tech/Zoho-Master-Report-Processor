import { useState } from 'react'
import { processRawFile } from '../api'
import FileDrop from './FileDrop'
import Alert from './Alert'
import { DEALS_QUARTERS, PIPELINE_PERIODS } from './QuarterSelector'

const PERIOD_OPTIONS_BY_TYPE = {
  deals: DEALS_QUARTERS,
  pipeline: PIPELINE_PERIODS,
}

export default function ProcessPanel({ reportType, reportLabel, quarter }) {
  const [file, setFile] = useState(null)
  const [status, setStatus] = useState('idle') // idle | working | done | error
  const [error, setError] = useState('')

  const periodLabel =
    PERIOD_OPTIONS_BY_TYPE[reportType]?.find((p) => p.id === quarter)?.label || quarter

  async function handleProcess() {
    if (!file) return
    setStatus('working')
    setError('')
    try {
      await processRawFile(reportType, quarter, file)
      setStatus('done')
    } catch (e) {
      setStatus('error')
      setError(e.message)
    }
  }

  function selectFile(f) {
    setFile(f)
    setStatus('idle')
    setError('')
  }

  return (
    <div className="panel">
      <div className="panel-icon process-icon">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none">
          <path d="M4 12h16M4 12l4-4M4 12l4 4" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
          <path d="M14 6h6v6M20 12v6h-6" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" opacity="0.4" />
        </svg>
      </div>
      <h2>Process {periodLabel}</h2>
      <p className="panel-hint">
        Upload the raw Excel file downloaded from Zoho CRM for {reportLabel} ({periodLabel})
        after applying the Close Date and Stage filters. The formula columns are added
        automatically.
      </p>

      <FileDrop file={file} onSelect={selectFile} />

      <div className="actions">
        <button disabled={!file || status === 'working'} onClick={handleProcess}>
          {status === 'working' ? (
            <>
              <span className="spinner" /> Processing…
            </>
          ) : (
            'Process & Download'
          )}
        </button>
        {file && status !== 'working' && (
          <button className="secondary" onClick={() => selectFile(null)}>
            Clear
          </button>
        )}
      </div>

      {status === 'done' && <Alert type="success">Done — file downloaded.</Alert>}
      {status === 'error' && <Alert type="error">{error}</Alert>}
    </div>
  )
}
