import { useState } from 'react'
import { downloadTemplate, uploadTemplate } from '../api'
import FileDrop from './FileDrop'
import Alert from './Alert'

export default function TemplatePanel({ reportType, reportLabel }) {
  const [file, setFile] = useState(null)
  const [status, setStatus] = useState('idle') // idle | working | done | error
  const [error, setError] = useState('')

  async function handleDownload() {
    setStatus('working')
    setError('')
    try {
      await downloadTemplate(reportType)
      setStatus('idle')
    } catch (e) {
      setStatus('error')
      setError(e.message)
    }
  }

  async function handleUpload() {
    if (!file) return
    setStatus('working')
    setError('')
    try {
      await uploadTemplate(reportType, file)
      setStatus('done')
      setFile(null)
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
      <div className="panel-icon template-icon">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none">
          <rect x="4" y="3" width="16" height="18" rx="2" stroke="currentColor" strokeWidth="1.8" />
          <path d="M8 8h8M8 12h8M8 16h5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
        </svg>
      </div>
      <h2>Manage {reportLabel} template</h2>
      <p className="panel-hint">
        {reportType === 'pipeline' ? (
          <>
            One shared template is used for both Pipeline views — the full pipeline and the
            &lt;40% subset have the same columns and formulas.
          </>
        ) : reportType === 'leads' ? (
          <>
            This template defines the New Leads sheet appended to the combined workbook,
            filtered to Lead Status = "New" within the current fiscal year.
          </>
        ) : (
          <>
            One shared template is used for all four quarters — Q1, Q2, Q3, and Q4 all have the
            same columns and formulas, only the close-date range differs.
          </>
        )}{' '}
        Add a header in row 2 and a formula in row 3 (e.g. <code>=E3-F3</code>) for a calculated
        column, or an example value for a column copied straight from the raw file. Any change
        here applies automatically the next time you process a report.
      </p>

      <button className="secondary block" onClick={handleDownload} disabled={status === 'working'}>
        Download current template
      </button>

      <div className="divider">then, once edited</div>

      <FileDrop file={file} onSelect={selectFile} label="Upload the edited template" />

      {file && (
        <div className="actions">
          <button disabled={status === 'working'} onClick={handleUpload}>
            {status === 'working' ? (
              <>
                <span className="spinner" /> Uploading…
              </>
            ) : (
              'Save template'
            )}
          </button>
          <button className="secondary" onClick={() => selectFile(null)} disabled={status === 'working'}>
            Clear
          </button>
        </div>
      )}

      {status === 'done' && <Alert type="success">Template updated.</Alert>}
      {status === 'error' && <Alert type="error">{error}</Alert>}
    </div>
  )
}
