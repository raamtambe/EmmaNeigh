import { useState, useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import HorseAnimation from '../components/HorseAnimation'
import ProgressBar from '../components/ProgressBar'

// Processing states
const STATE = {
  IDLE: 'idle',
  PROCESSING: 'processing',
  COMPLETE: 'complete',
  ERROR: 'error',
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
    <div className="bg-gray-50 rounded-xl p-4">
      <label className="block text-sm font-semibold text-gray-700 mb-1">
        {label}
      </label>
      <p className="text-xs text-gray-500 mb-3">{description}</p>

      {file ? (
        <div className="flex items-center justify-between bg-white rounded-lg p-3 border border-gray-200">
          <div className="flex items-center gap-2">
            <span className="text-2xl">📄</span>
            <span className="text-sm text-gray-700 truncate max-w-xs">{file.name}</span>
          </div>
          <button
            onClick={onClear}
            className="text-red-500 hover:text-red-700 text-sm"
          >
            Remove
          </button>
        </div>
      ) : (
        <button
          onClick={handleSelect}
          className="w-full py-3 border-2 border-dashed border-gray-300 rounded-lg
                   text-gray-500 hover:border-blue-400 hover:text-blue-500 hover:bg-blue-50
                   transition-colors"
        >
          Click to select PDF
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

  // Set up progress listener
  useEffect(() => {
    if (!window.api) return

    const cleanup = window.api.onProgress((data) => {
      if (data.type === 'progress') {
        setProgress({
          percent: data.percent,
          message: data.message,
        })
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
    <div className="min-h-screen flex flex-col">
      {/* Header */}
      <header className="bg-white shadow-sm">
        <div className="max-w-4xl mx-auto px-6 py-4 flex items-center gap-4">
          <button
            onClick={() => navigate('/')}
            className="text-gray-500 hover:text-gray-700 transition-colors"
          >
            <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
            </svg>
          </button>
          <h1 className="text-2xl font-bold text-emma-navy">Create Execution Version</h1>
        </div>
      </header>

      {/* Main content */}
      <main className="flex-1 max-w-4xl mx-auto w-full px-6 py-8">
        {/* Idle State - File Selection */}
        {state === STATE.IDLE && (
          <div className="space-y-6 animate-fade-in">
            <div className="bg-white rounded-2xl shadow-lg p-8">
              <h2 className="text-xl font-semibold text-gray-800 mb-2">
                Merge Signed Pages
              </h2>
              <p className="text-gray-600 mb-6">
                Combine signed pages from DocuSign back into the original document to create
                the final execution version. DocuSign PDFs will be automatically unlocked.
              </p>

              <div className="space-y-4">
                {/* Step 1: Original PDF */}
                <FileSelector
                  label="Step 1: Original PDF"
                  description="The clean document without signature pages"
                  file={originalFile}
                  onSelect={setOriginalFile}
                  onClear={() => setOriginalFile(null)}
                />

                {/* Step 2: Signed PDF */}
                <FileSelector
                  label="Step 2: Signed PDF (from DocuSign)"
                  description="The signed document received from DocuSign"
                  file={signedFile}
                  onSelect={setSignedFile}
                  onClear={() => setSignedFile(null)}
                />

                {/* Step 3: Insertion Point */}
                <div className="bg-gray-50 rounded-xl p-4">
                  <label className="block text-sm font-semibold text-gray-700 mb-1">
                    Step 3: Insert signed pages after page
                  </label>
                  <p className="text-xs text-gray-500 mb-3">
                    Leave blank to insert at the end of the document
                  </p>
                  <input
                    type="number"
                    min="0"
                    value={insertAfter}
                    onChange={(e) => setInsertAfter(e.target.value)}
                    placeholder="e.g., 85"
                    className="w-32 px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  />
                </div>
              </div>

              {canProcess && (
                <div className="mt-6 flex justify-center">
                  <button
                    onClick={handleProcess}
                    className="px-8 py-3 bg-blue-600 text-white font-semibold rounded-xl
                             hover:bg-blue-700 transition-colors shadow-lg hover:shadow-xl"
                  >
                    Create Execution Version
                  </button>
                </div>
              )}
            </div>

            {/* Info card */}
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-6">
              <h3 className="font-semibold text-amber-800 mb-2">About DocuSign PDFs</h3>
              <p className="text-amber-700 text-sm">
                DocuSign returns PDFs with permission restrictions that prevent editing.
                This tool automatically removes those restrictions so the signed pages
                can be merged into your original document.
              </p>
            </div>
          </div>
        )}

        {/* Processing State */}
        {state === STATE.PROCESSING && (
          <div className="bg-white rounded-2xl shadow-lg p-8 animate-fade-in">
            <HorseAnimation
              statusMessage={progress.message || 'Processing...'}
              isRunning={true}
            />

            <div className="mt-8 max-w-md mx-auto">
              <ProgressBar percent={progress.percent} />
            </div>

            <p className="text-center text-gray-500 mt-6 text-sm">
              Creating your execution version...
            </p>
          </div>
        )}

        {/* Complete State */}
        {state === STATE.COMPLETE && result && (
          <div className="bg-white rounded-2xl shadow-lg p-8 animate-fade-in">
            <div className="text-center mb-8">
              <div className="text-6xl mb-4">✅</div>
              <h2 className="text-2xl font-bold text-green-600 mb-2">Complete!</h2>
              <p className="text-gray-600">
                Execution version created successfully
              </p>
            </div>

            {/* Results info */}
            <div className="bg-gray-50 rounded-xl p-4 mb-6">
              <h3 className="font-semibold text-gray-700 mb-2">Document Details:</h3>
              <ul className="space-y-1 text-sm">
                <li className="flex justify-between">
                  <span className="text-gray-600">Original pages:</span>
                  <span className="text-gray-800">{result.originalPages}</span>
                </li>
                <li className="flex justify-between">
                  <span className="text-gray-600">Signed pages added:</span>
                  <span className="text-gray-800">{result.signedPages}</span>
                </li>
                <li className="flex justify-between font-semibold">
                  <span className="text-gray-700">Total pages:</span>
                  <span className="text-gray-900">{result.totalPages}</span>
                </li>
              </ul>
            </div>

            {/* Action buttons */}
            <div className="flex justify-center gap-4">
              <button
                onClick={handleDownload}
                className="px-8 py-3 bg-green-600 text-white font-semibold rounded-xl
                         hover:bg-green-700 transition-colors shadow-lg hover:shadow-xl
                         flex items-center gap-2"
              >
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                        d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                </svg>
                Download PDF
              </button>

              <button
                onClick={handleReset}
                className="px-8 py-3 bg-gray-100 text-gray-700 font-semibold rounded-xl
                         hover:bg-gray-200 transition-colors"
              >
                Start Over
              </button>
            </div>
          </div>
        )}

        {/* Error State */}
        {state === STATE.ERROR && (
          <div className="bg-white rounded-2xl shadow-lg p-8 animate-fade-in">
            <div className="text-center mb-6">
              <div className="text-6xl mb-4">❌</div>
              <h2 className="text-2xl font-bold text-red-600 mb-2">Error</h2>
              <p className="text-gray-600">{error}</p>
            </div>

            <div className="flex justify-center">
              <button
                onClick={handleReset}
                className="px-8 py-3 bg-blue-600 text-white font-semibold rounded-xl
                         hover:bg-blue-700 transition-colors"
              >
                Try Again
              </button>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="py-4 text-center text-sm text-gray-500">
        All processing is done locally. No data leaves your machine.
      </footer>
    </div>
  )
}
