import { useCallback } from 'react'
import { useDropzone } from 'react-dropzone'

function FolderIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path strokeLinecap="round" strokeLinejoin="round" d="M2.25 12.75V12A2.25 2.25 0 014.5 9.75h15A2.25 2.25 0 0121.75 12v.75m-8.69-6.44l-2.12-2.12a1.5 1.5 0 00-1.061-.44H4.5A2.25 2.25 0 002.25 6v12a2.25 2.25 0 002.25 2.25h15A2.25 2.25 0 0021.75 18V9a2.25 2.25 0 00-2.25-2.25h-5.379a1.5 1.5 0 01-1.06-.44z" />
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

export default function FileUpload({
  onFilesSelected,
  accept = { 'application/pdf': ['.pdf'] },
  multiple = true,
  disabled = false,
  selectedFiles = [],
}) {
  const onDrop = useCallback((acceptedFiles) => {
    if (acceptedFiles.length > 0) {
      onFilesSelected(acceptedFiles)
    }
  }, [onFilesSelected])

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept,
    multiple,
    disabled,
  })

  const handleBrowse = async () => {
    if (window.api) {
      const files = await window.api.selectFiles()
      if (files && files.length > 0) {
        const fileObjects = files.map(path => ({
          path,
          name: path.split(/[/\\]/).pop(),
        }))
        onFilesSelected(fileObjects)
      }
    }
  }

  return (
    <div className="w-full">
      <div
        {...getRootProps()}
        className={`
          border border-dashed rounded-lg p-6 text-center cursor-pointer transition-all duration-200
          ${isDragActive
            ? 'border-slate-400 bg-slate-50'
            : disabled
              ? 'border-slate-200 bg-slate-50 cursor-not-allowed'
              : 'border-slate-300 bg-white hover:border-slate-400 hover:bg-slate-50'
          }
        `}
      >
        <input {...getInputProps()} />

        {selectedFiles.length > 0 ? (
          <div className="space-y-3">
            <DocumentIcon className="w-8 h-8 text-slate-400 mx-auto" />
            <div>
              <p className="text-sm font-medium text-slate-700">
                {selectedFiles.length} file{selectedFiles.length !== 1 ? 's' : ''} selected
              </p>
              <div className="mt-2 max-h-24 overflow-y-auto">
                {selectedFiles.slice(0, 5).map((file, idx) => (
                  <p key={idx} className="text-xs text-slate-500 truncate">
                    {file.name || file.path?.split(/[/\\]/).pop()}
                  </p>
                ))}
                {selectedFiles.length > 5 && (
                  <p className="text-xs text-slate-400 mt-1">
                    +{selectedFiles.length - 5} more
                  </p>
                )}
              </div>
            </div>
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation()
                onFilesSelected([])
              }}
              className="text-xs text-slate-500 hover:text-slate-700"
            >
              Clear
            </button>
          </div>
        ) : (
          <div className="space-y-3">
            <FolderIcon className="w-8 h-8 text-slate-400 mx-auto" />
            <div>
              <p className="text-sm text-slate-600">Drop PDF files here</p>
              <p className="text-xs text-slate-400 mt-1">or click to browse</p>
            </div>

            {window.api && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation()
                  handleBrowse()
                }}
                className="px-3 py-1.5 text-xs font-medium text-slate-600 bg-slate-100 rounded-md hover:bg-slate-200 transition-colors"
              >
                Browse Files
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  )
}
