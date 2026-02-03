import { useState, useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import FileUpload from '../components/FileUpload'
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

export default function SignaturePackets() {
  const navigate = useNavigate()
  const [state, setState] = useState(STATE.IDLE)
  const [selectedFiles, setSelectedFiles] = useState([])
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

  const handleFilesSelected = (files) => {
    setSelectedFiles(files)
    setError(null)
  }

  const handleProcess = async () => {
    if (selectedFiles.length === 0) return

    setState(STATE.PROCESSING)
    setProgress({ percent: 0, message: 'Starting...' })
    setError(null)

    try {
      const filePaths = selectedFiles.map(f => f.path || f.name)
      const response = await window.api.processSignaturePackets(filePaths)

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
    if (!result?.zipPath) return
    const savePath = await window.api.saveFile('EmmaNeigh-Signature-Packets.zip')
    if (savePath) {
      await window.api.copyFile(result.zipPath, savePath)
    }
  }

  const handleReset = () => {
    setState(STATE.IDLE)
    setSelectedFiles([])
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
          <h1 className="text-lg font-semibold text-slate-900">Create Signature Packets</h1>
        </div>
      </header>

      {/* Main content */}
      <main className="flex-1 max-w-3xl mx-auto w-full px-6 py-8">
        {/* Idle State */}
        {state === STATE.IDLE && (
          <div className="space-y-6">
            <div className="bg-white rounded-lg border border-slate-200 p-6">
              <h2 className="text-base font-medium text-slate-900 mb-1">Select PDF files</h2>
              <p className="text-sm text-slate-500 mb-6">
                Choose documents containing signature pages. The tool will identify signers and create individual packets.
              </p>

              <FileUpload
                onFilesSelected={handleFilesSelected}
                selectedFiles={selectedFiles}
              />

              {selectedFiles.length > 0 && (
                <div className="mt-6 flex justify-end">
                  <button
                    onClick={handleProcess}
                    className="px-4 py-2 bg-slate-900 text-white text-sm font-medium rounded-md hover:bg-slate-800 transition-colors"
                  >
                    Create Packets
                  </button>
                </div>
              )}
            </div>

            {/* Info */}
            <div className="bg-slate-100 rounded-lg p-4">
              <h3 className="text-sm font-medium text-slate-700 mb-2">How it works</h3>
              <ol className="text-sm text-slate-600 space-y-1 list-decimal list-inside">
                <li>Upload PDF documents containing signature pages</li>
                <li>The tool scans for "BY:" and "Name:" fields to identify signers</li>
                <li>Pages are grouped by signer into individual packets</li>
                <li>Download a ZIP file with all packets</li>
              </ol>
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
              <p className="text-sm text-slate-500">
                Created {result.packetsCreated || 0} signature packet{result.packetsCreated !== 1 ? 's' : ''}
              </p>
            </div>

            {result.packets && result.packets.length > 0 && (
              <div className="bg-slate-50 rounded-md p-4 mb-6 max-h-40 overflow-y-auto">
                <ul className="space-y-1">
                  {result.packets.map((packet, idx) => (
                    <li key={idx} className="flex justify-between text-sm">
                      <span className="text-slate-700">{packet.name}</span>
                      <span className="text-slate-400">{packet.pages} pg</span>
                    </li>
                  ))}
                </ul>
              </div>
            )}

            <div className="flex justify-center gap-3">
              <button
                onClick={handleDownload}
                className="px-4 py-2 bg-slate-900 text-white text-sm font-medium rounded-md hover:bg-slate-800 transition-colors flex items-center gap-2"
              >
                <DownloadIcon className="w-4 h-4" />
                Download ZIP
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
