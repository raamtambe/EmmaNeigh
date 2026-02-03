import { useState, useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import HorseAnimation from '../components/HorseAnimation'
import ProgressBar from '../components/ProgressBar'

const STATE = {
  IDLE: 'idle',
  PROCESSING: 'processing',
  COMPLETE: 'complete',
  ERROR: 'error',
}

// Icons
function BackIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 19.5L8.25 12l7.5-7.5" />
    </svg>
  )
}

function DocumentIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
    </svg>
  )
}

function DownloadIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3" />
    </svg>
  )
}

function CheckIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <path strokeLinecap="round" strokeLinejoin="round" d="M4.5 12.75l6 6 9-13.5" />
    </svg>
  )
}

function AlertIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v3.75m9-.75a9 9 0 11-18 0 9 9 0 0118 0zm-9 3.75h.008v.008H12v-.008z" />
    </svg>
  )
}

function XIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
    </svg>
  )
}

function FileSelector({ label, description, file, onSelect, onClear }) {
  const handleSelect = async () => {
    if (window.api) {
      const files = await window.api.selectFiles()
      if (files && files.length > 0) {
        onSelect({
          path: files[0],
          name: files[0].split(/[/\\]/).pop(),
        })
      }
    }
  }

  return (
    <div className="space-y-2">
      <label className="block text-sm font-medium text-slate-700">{label}</label>
      <p className="text-xs text-slate-500">{description}</p>

      {file ? (
        <div className="flex items-center justify-between bg-slate-50 rounded-md px-3 py-2 border border-slate-200">
          <div className="flex items-center gap-2 min-w-0">
            <DocumentIcon className="w-4 h-4 text-slate-400 flex-shrink-0" />
            <span className="text-sm text-slate-700 truncate">{file.name}</span>
          </div>
          <button
            onClick={onClear}
            className="p-1 text-slate-400 hover:text-slate-600 flex-shrink-0"
          >
            <XIcon className="w-4 h-4" />
          </button>
        </div>
      ) : (
        <button
          onClick={handleSelect}
          className="w-full py-2.5 border border-dashed border-slate-300 rounded-md text-sm text-slate-500 hover:border-slate-400 hover:text-slate-600 hover:bg-slate-50 transition-colors"
        >
          Select PDF
        </button>
      )}
    </div>
  )
}

export default function ExecutionVersion() {
  const navigate = useNavigate()
  const [state, setState] = useState(STATE.IDLE)
  const [originalFile, setOriginalFile] = useState(null)
  const [signedFile, setSignedFile] = useState(null)
  const [insertAfter, setInsertAfter] = useState('')
  const [progress, setProgress] = useState({ percent: 0, message: '' })
  const [result, setResult] = useState(null)
  const [error, setError] = useState(null)

  useEffect(() => {
    if (!window.api) return
    const cleanup = window.api.onProgress((data) => {
      if (data.type === 'progress') {
        setProgress({ percent: data.percent, message: data.message })
      }
    })
    return cleanup
  }, [])

  const canProcess = originalFile && signedFile

  const handleProcess = async () => {
    if (!canProcess) return

    setState(STATE.PROCESSING)
    setProgress({ percent: 0, message: 'Starting...' })
    setError(null)

    try {
      const pageNum = insertAfter === '' ? -1 : parseInt(insertAfter, 10)
      const response = await window.api.createExecutionVersion(
        originalFile.path,
        signedFile.path,
        pageNum
      )

      if (response.success) {
        setResult(response)
        setState(STATE.COMPLETE)
      } else {
        throw new Error(response.error || 'Processing failed')
      }
    } catch (err) {
      setError(err.message || 'An error occurred during processing')
      setState(STATE.ERROR)
    }
  }

  const handleDownload = async () => {
    if (!result?.outputPath) return
    const defaultName = result.outputFilename || 'Execution Version.pdf'
    const savePath = await window.api.saveFile(defaultName)
    if (savePath) {
      await window.api.copyFile(result.outputPath, savePath)
    }
  }

  const handleReset = () => {
    setState(STATE.IDLE)
    setOriginalFile(null)
    setSignedFile(null)
    setInsertAfter('')
    setProgress({ percent: 0, message: '' })
    setResult(null)
    setError(null)
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      {/* Header */}
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-3xl mx-auto px-6 py-4 flex items-center gap-3">
          <button
            onClick={() => navigate('/')}
            className="p-1.5 -ml-1.5 text-slate-400 hover:text-slate-600 transition-colors rounded-md hover:bg-slate-100"
          >
            <BackIcon className="w-5 h-5" />
          </button>
          <h1 className="text-lg font-semibold text-slate-900">Create Execution Version</h1>
        </div>
      </header>

      {/* Main content */}
      <main className="flex-1 max-w-3xl mx-auto w-full px-6 py-8">
        {/* Idle State */}
        {state === STATE.IDLE && (
          <div className="space-y-6">
            <div className="bg-white rounded-lg border border-slate-200 p-6">
              <h2 className="text-base font-medium text-slate-900 mb-1">Merge signed pages</h2>
              <p className="text-sm text-slate-500 mb-6">
                Combine signed pages from DocuSign back into the original document.
              </p>

              <div className="space-y-5">
                <FileSelector
                  label="Original PDF"
                  description="The clean document without signature pages"
                  file={originalFile}
                  onSelect={setOriginalFile}
                  onClear={() => setOriginalFile(null)}
                />

                <FileSelector
                  label="Signed PDF"
                  description="The signed document from DocuSign"
                  file={signedFile}
                  onSelect={setSignedFile}
                  onClear={() => setSignedFile(null)}
                />

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">
                    Insert after page
                  </label>
                  <p className="text-xs text-slate-500">
                    Leave blank to insert at the end
                  </p>
                  <input
                    type="number"
                    min="0"
                    value={insertAfter}
                    onChange={(e) => setInsertAfter(e.target.value)}
                    placeholder="e.g., 85"
                    className="w-24 px-3 py-1.5 text-sm border border-slate-300 rounded-md focus:ring-1 focus:ring-slate-400 focus:border-slate-400 outline-none"
                  />
                </div>
              </div>

              {canProcess && (
                <div className="mt-6 flex justify-end">
                  <button
                    onClick={handleProcess}
                    className="px-4 py-2 bg-slate-900 text-white text-sm font-medium rounded-md hover:bg-slate-800 transition-colors"
                  >
                    Create Execution Version
                  </button>
                </div>
              )}
            </div>

            {/* Info */}
            <div className="bg-slate-100 rounded-lg p-4">
              <h3 className="text-sm font-medium text-slate-700 mb-2">About DocuSign PDFs</h3>
              <p className="text-sm text-slate-600">
                DocuSign returns PDFs with permission restrictions. This tool automatically removes those restrictions so pages can be merged.
              </p>
            </div>
          </div>
        )}

        {/* Processing State */}
        {state === STATE.PROCESSING && (
          <div className="bg-white rounded-lg border border-slate-200 p-8">
            <HorseAnimation
              statusMessage={progress.message || 'Processing...'}
              isRunning={true}
            />
            <div className="mt-6 max-w-sm mx-auto">
              <ProgressBar percent={progress.percent} />
            </div>
          </div>
        )}

        {/* Complete State */}
        {state === STATE.COMPLETE && result && (
          <div className="bg-white rounded-lg border border-slate-200 p-8">
            <div className="text-center mb-6">
              <div className="w-12 h-12 bg-emerald-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <CheckIcon className="w-6 h-6 text-emerald-600" />
              </div>
              <h2 className="text-lg font-semibold text-slate-900 mb-1">Complete</h2>
              <p className="text-sm text-slate-500">Execution version created</p>
            </div>

            <div className="bg-slate-50 rounded-md p-4 mb-6">
              <dl className="space-y-1 text-sm">
                <div className="flex justify-between">
                  <dt className="text-slate-500">Original pages</dt>
                  <dd className="text-slate-700">{result.originalPages}</dd>
                </div>
                <div className="flex justify-between">
                  <dt className="text-slate-500">Signed pages added</dt>
                  <dd className="text-slate-700">{result.signedPages}</dd>
                </div>
                <div className="flex justify-between font-medium">
                  <dt className="text-slate-700">Total pages</dt>
                  <dd className="text-slate-900">{result.totalPages}</dd>
                </div>
              </dl>
            </div>

            <div className="flex justify-center gap-3">
              <button
                onClick={handleDownload}
                className="px-4 py-2 bg-slate-900 text-white text-sm font-medium rounded-md hover:bg-slate-800 transition-colors flex items-center gap-2"
              >
                <DownloadIcon className="w-4 h-4" />
                Download PDF
              </button>
              <button
                onClick={handleReset}
                className="px-4 py-2 bg-slate-100 text-slate-700 text-sm font-medium rounded-md hover:bg-slate-200 transition-colors"
              >
                Start Over
              </button>
            </div>
          </div>
        )}

        {/* Error State */}
        {state === STATE.ERROR && (
          <div className="bg-white rounded-lg border border-slate-200 p-8">
            <div className="text-center mb-6">
              <div className="w-12 h-12 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <AlertIcon className="w-6 h-6 text-red-600" />
              </div>
              <h2 className="text-lg font-semibold text-slate-900 mb-1">Error</h2>
              <p className="text-sm text-slate-500">{error}</p>
            </div>
            <div className="flex justify-center">
              <button
                onClick={handleReset}
                className="px-4 py-2 bg-slate-900 text-white text-sm font-medium rounded-md hover:bg-slate-800 transition-colors"
              >
                Try Again
              </button>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-200 bg-white">
        <div className="max-w-3xl mx-auto px-6 py-4">
          <p className="text-xs text-slate-400">
            All processing happens locally. No data leaves your machine.
          </p>
        </div>
      </footer>
    </div>
  )
}
