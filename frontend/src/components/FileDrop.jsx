import { useRef, useState } from 'react'

export default function FileDrop({ file, onSelect, accept = '.xlsx,.xlsm', label }) {
  const inputRef = useRef(null)
  const [dragActive, setDragActive] = useState(false)

  function handleFiles(fileList) {
    const picked = fileList?.[0]
    if (picked) onSelect(picked)
  }

  return (
    <div
      className={`dropzone ${dragActive ? 'drag-active' : ''} ${file ? 'has-file' : ''}`}
      onClick={() => inputRef.current?.click()}
      onDragOver={(e) => {
        e.preventDefault()
        setDragActive(true)
      }}
      onDragLeave={() => setDragActive(false)}
      onDrop={(e) => {
        e.preventDefault()
        setDragActive(false)
        handleFiles(e.dataTransfer.files)
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        hidden
        onChange={(e) => handleFiles(e.target.files)}
      />
      <svg width="28" height="28" viewBox="0 0 24 24" fill="none" className="dropzone-icon">
        <path
          d="M12 16V4M12 4L7 9M12 4l5 5"
          stroke="currentColor"
          strokeWidth="1.8"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
        <path
          d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2"
          stroke="currentColor"
          strokeWidth="1.8"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
      {file ? (
        <div className="dropzone-file">
          <span className="file-chip">{file.name}</span>
          <span className="dropzone-sub">Click or drop to replace</span>
        </div>
      ) : (
        <div>
          <div className="dropzone-label">{label || 'Click to browse, or drop a file here'}</div>
          <div className="dropzone-sub">.xlsx or .xlsm</div>
        </div>
      )}
    </div>
  )
}
