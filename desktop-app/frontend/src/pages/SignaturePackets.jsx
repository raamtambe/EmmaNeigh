import { useState, useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import FileUpload from '../components/FileUpload'
import HorseAnimation from '../components/HorseAnimation'
import ProgressBar from '../components/ProgressBar'

// Processing states
const STATE = {
  IDLE: 'idle',
  PROCESSING: 'processing',
  COMPLETE: 'complete',
  ERROR: 'error',
}

export default function SignaturePackets() {
  const navigate = useNavigate()
  const [state, setState] = useState(STATE.IDLE)
  const [selectedFiles, setSelectedFiles] = useState([])
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
      // Get file paths
      const filePaths = selectedFiles.map(f => f.path || f.name)

      // Call the Electron API
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
      // Could show a success notification here
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
          <h1 className="text-2xl font-bold text-emma-navy">Create Signature Packets</h1>
        </div>
      </header>

      {/* Main content */}
      <main className="flex-1 max-w-4xl mx-auto w-full px-6 py-8">
        {/* Idle State - File Selection */}
        {state === STATE.IDLE && (
          <div className="space-y-6 animate-fade-in">
            <div className="bg-white rounded-2xl shadow-lg p-8">
              <h2 className="text-xl font-semibold text-gray-800 mb-4">
                Select PDF Files
              </h2>
              <p className="text-gray-600 mb-6">
                Choose the PDF documents containing signature pages. The tool will scan each document,
                identify signature pages, and organize them by signer.
              </p>

              <FileUpload
                onFilesSelected={handleFilesSelected}
                selectedFiles={selectedFiles}
                title="Drag & drop PDF files here"
                subtitle="or click to browse your files"
              />

              {selectedFiles.length > 0 && (
                <div className="mt-6 flex justify-center">
                  <button
                    onClick={handleProcess}
                    className="px-8 py-3 bg-blue-600 text-white font-semibold rounded-xl
                             hover:bg-blue-700 transition-colors shadow-lg hover:shadow-xl"
                  >
                    Create Signature Packets
                  </button>
                </div>
              )}
            </div>

            {/* Info card */}
            <div className="bg-blue-50 border border-blue-200 rounded-xl p-6">
              <h3 className="font-semibold text-blue-800 mb-2">How it works</h3>
              <ul className="text-blue-700 text-sm space-y-1">
                <li>1. The tool scans each PDF for signature pages</li>
                <li>2. It identifies signers by looking for "BY:" and "Name:" fields</li>
                <li>3. Pages are grouped by signer into individual packets</li>
                <li>4. You'll receive a ZIP file with all packets and an Excel index</li>
              </ul>
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
              Please wait while your documents are being processed...
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
                Created {result.packetsCreated || 0} signature packet{result.packetsCreated !== 1 ? 's' : ''}
              </p>
            </div>

            {/* Results list */}
            {result.packets && result.packets.length > 0 && (
              <div className="bg-gray-50 rounded-xl p-4 mb-6 max-h-48 overflow-y-auto">
                <h3 className="font-semibold text-gray-700 mb-2">Packets Created:</h3>
                <ul className="space-y-1">
                  {result.packets.map((packet, idx) => (
                    <li key={idx} className="flex justify-between text-sm">
                      <span className="text-gray-700">{packet.name}</span>
                      <span className="text-gray-500">{packet.pages} page{packet.pages !== 1 ? 's' : ''}</span>
                    </li>
                  ))}
                </ul>
              </div>
            )}

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
                Download ZIP
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
