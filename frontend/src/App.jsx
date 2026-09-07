import { useState } from 'react'
import ReportTypeSelector from './components/ReportTypeSelector'
import ProcessPanel from './components/ProcessPanel'
import TemplatePanel from './components/TemplatePanel'

const REPORT_LABELS = {
  deals: 'Deals',
  pipeline: 'Pipeline',
}

const DEFAULT_PERIOD = {
  deals: 'Q1',
  pipeline: 'FULL',
}

export default function App() {
  const [reportType, setReportType] = useState('deals')
  const [quarter, setQuarter] = useState(DEFAULT_PERIOD.deals)
  const reportLabel = REPORT_LABELS[reportType]

  function selectReportType(type) {
    setReportType(type)
    setQuarter(DEFAULT_PERIOD[type])
  }

  return (
    <div className="app-dashboard">
      <aside className="sidebar-col">
        <header>
          <div className="brand">
            <div className="brand-mark">Z</div>
            <div>
              <h1>Zoho Master Report Processor</h1>
              <p className="subtitle">Add the quarterly formula columns automatically.</p>
            </div>
          </div>
        </header>

        <section className="nav-section">
          <h3 className="section-label">Report Type &amp; Quarter</h3>
          <ReportTypeSelector
            selected={reportType}
            onSelect={selectReportType}
            quarter={quarter}
            onQuarterSelect={setQuarter}
          />
        </section>
      </aside>

      <main className="main-workspace">
          <section>
            <h3 className="section-label">1. Process a report</h3>
            <ProcessPanel reportType={reportType} reportLabel={reportLabel} quarter={quarter} />
          </section>

          <section>
            <h3 className="section-label">2. Template (shared)</h3>
            <TemplatePanel reportType={reportType} reportLabel={reportLabel} />
          </section>
      </main>
    </div>
  )
}
