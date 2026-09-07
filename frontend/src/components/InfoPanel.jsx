const STEPS = [
  { title: 'Filter in Zoho CRM', detail: 'Master Report → filter by Close Date (the quarter) and Stage.' },
  { title: 'Download the raw export', detail: 'Save the filtered result as an .xlsx file.' },
  { title: 'Upload it here', detail: 'Pick the report type and quarter tab, then drop the file in.' },
  { title: 'Get the finished report', detail: 'The formula columns are added automatically — download and you’re done.' },
]

const FORMULAS_BY_TYPE = {
  deals: [
    { name: 'Balance Disbursement Amount', formula: 'Sanctioned − Disbursed' },
    { name: 'Diff. of Sanction & Disbursed Amount', formula: 'Sanctioned − Disbursed' },
    { name: 'Diff. of Requested & Sanctioned Amount', formula: 'Requested − Sanctioned' },
    { name: '% Requested vs Sanctioned', formula: '(Sanctioned − Requested) / Requested' },
    { name: '% Sanctioned vs Disbursed', formula: '(Disbursed − Sanctioned) / Sanctioned' },
  ],
  pipeline: [
    { name: 'Diff. of Requested vs Login Amount', formula: 'Requested − Login' },
    { name: '% Requested vs Login', formula: 'Login / Requested' },
  ],
}

export default function InfoPanel({ reportType }) {
  const formulas = FORMULAS_BY_TYPE[reportType] || []

  return (
    <aside className="info-col">
      <div className="info-card">
        <h3 className="info-title">How it works</h3>
        <ol className="steps">
          {STEPS.map((s, i) => (
            <li key={i}>
              <span className="step-num">{i + 1}</span>
              <div className="step-content">
                <div className="step-title">{s.title}</div>
                <div className="step-detail">{s.detail}</div>
              </div>
            </li>
          ))}
        </ol>
      </div>

      <div className="info-card">
        <h3 className="info-title">The formula columns</h3>
        <ul className="formula-list">
          {formulas.map((f) => (
            <li key={f.name}>
              <div className="formula-name">{f.name}</div>
              <div className="formula-expr">{f.formula}</div>
            </li>
          ))}
        </ul>
        {reportType === 'pipeline' && (
          <p className="panel-hint" style={{ marginTop: 12, marginBottom: 0 }}>
            Left blank for a deal until its Login Amount is known.
          </p>
        )}
      </div>
    </aside>
  )
}
