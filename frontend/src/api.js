// In production (Netlify), set VITE_API_BASE to the deployed Render backend's
// URL, e.g. https://zoho-report-backend.onrender.com/api. Falls back to the
// local dev backend when that env var isn't set.
const API_BASE = import.meta.env.VITE_API_BASE || 'http://localhost:8000/api'

async function extractError(res) {
  try {
    const data = await res.json()
    return data.detail || res.statusText
  } catch {
    return res.statusText
  }
}

async function apiFetch(url, options) {
  try {
    return await fetch(url, options)
  } catch {
    throw new Error(
      'Could not reach the server. Make sure the backend is running at localhost:8000.'
    )
  }
}

function filenameFromDisposition(res, fallback) {
  const disposition = res.headers.get('Content-Disposition') || ''
  const match = disposition.match(/filename="?([^"]+)"?/)
  return match ? match[1] : fallback
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  a.remove()
  URL.revokeObjectURL(url)
}

export async function downloadTemplate(reportType) {
  const res = await apiFetch(`${API_BASE}/templates/${reportType}`)
  if (!res.ok) throw new Error(await extractError(res))
  const blob = await res.blob()
  downloadBlob(blob, filenameFromDisposition(res, `${reportType}_template.xlsx`))
}

export async function uploadTemplate(reportType, file) {
  const form = new FormData()
  form.append('file', file)
  const res = await apiFetch(`${API_BASE}/templates/${reportType}`, {
    method: 'POST',
    body: form,
  })
  if (!res.ok) throw new Error(await extractError(res))
  return res.json()
}

export async function processMasterReport(file, leadsFile, meetingsFile) {
  const form = new FormData()
  form.append('file', file)
  form.append('leads_file', leadsFile)
  form.append('meetings_file', meetingsFile)
  const res = await apiFetch(`${API_BASE}/process/master`, {
    method: 'POST',
    body: form,
  })
  if (!res.ok) throw new Error(await extractError(res))
  const blob = await res.blob()
  downloadBlob(blob, filenameFromDisposition(res, `MIS ${new Date().toISOString().slice(0, 10)}.xlsx`))
}
