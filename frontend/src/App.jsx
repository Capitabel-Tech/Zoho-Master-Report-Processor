import TemplatePanel from './components/TemplatePanel'
import MasterReportPanel from './components/MasterReportPanel'

export default function App() {
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
      </aside>

      <main className="main-workspace">
        <section>
          <h3 className="section-label">1. Process the Master Report</h3>
          <MasterReportPanel />
        </section>

        <section>
          <h3 className="section-label">2. Manage Deals template</h3>
          <TemplatePanel reportType="deals" reportLabel="Deals" />
        </section>

        <section>
          <h3 className="section-label">3. Manage Pipeline template</h3>
          <TemplatePanel reportType="pipeline" reportLabel="Pipeline" />
        </section>

        <section>
          <h3 className="section-label">4. Manage Leads template</h3>
          <TemplatePanel reportType="leads" reportLabel="Leads" />
        </section>
      </main>
    </div>
  )
}
