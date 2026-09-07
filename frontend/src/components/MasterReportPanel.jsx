import { useState } from 'react'
import { processMasterReport } from '../api'
import FileDrop from './FileDrop'
import Alert from './Alert'

export default function MasterReportPanel() {
  const [file, setFile] = useState(null)
  const [leadsFile, setLeadsFile] = useState(null)
  const [status, setStatus] = useState('idle') // idle | working | done | error
  const [error, setError] = useState('')

  const ready = file && leadsFile

  async function handleProcess() {
    if (!ready) return
    setStatus('working')
    setError('')
    try {
      await processMasterReport(file, leadsFile)
      setStatus('done')
    } catch (e) {
      setStatus('error')
      setError(e.message)
    }
  }

  function clearAll() {
    setFile(null)
    setLeadsFile(null)
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
      <h2>Process Master Report</h2>
      <p className="panel-hint">
        Upload both <strong>unfiltered</strong> exports from Zoho CRM — no need to apply
        any filters first. Based on today's date, the app automatically builds one
        workbook with a Deals sheet per completed quarter this fiscal year, a "Till Date"
        sheet for the quarter in progress, two Pipeline sheets (split by stage), and a New
        Leads sheet for this fiscal year.
      </p>

      <p className="panel-hint" style={{ marginBottom: 6, fontWeight: 600 }}>
        1. Master Report (Deals &amp; Pipeline data)
      </p>
      <FileDrop file={file} onSelect={(f) => { setFile(f); setStatus('idle'); setError('') }} />

      <p className="panel-hint" style={{ margin: '16px 0 6px', fontWeight: 600 }}>
        2. Leads Module report
      </p>
      <FileDrop
        file={leadsFile}
        onSelect={(f) => { setLeadsFile(f); setStatus('idle'); setError('') }}
      />

      <div className="actions">
        <button disabled={!ready || status === 'working'} onClick={handleProcess}>
          {status === 'working' ? (
            <>
              <span className="spinner" /> Processing…
            </>
          ) : (
            'Process & Download'
          )}
        </button>
        {(file || leadsFile) && status !== 'working' && (
          <button className="secondary" onClick={clearAll}>
            Clear
          </button>
        )}
      </div>

      {status === 'done' && <Alert type="success">Done — combined workbook downloaded.</Alert>}
      {status === 'error' && <Alert type="error">{error}</Alert>}
    </div>
  )
}
