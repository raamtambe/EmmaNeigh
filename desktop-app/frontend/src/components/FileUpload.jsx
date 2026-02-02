import { useCallback } from 'react'
import { useDropzone } from 'react-dropzone'

export default function FileUpload({
  onFilesSelected,
  accept = { 'application/pdf': ['.pdf'] },
  multiple = true,
  disabled = false,
  selectedFiles = [],
  title = 'Drag & drop PDF files here',
  subtitle = 'or click to browse',
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
        // Convert paths to file-like objects
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
          border-2 border-dashed rounded-xl p-8 text-center cursor-pointer
          transition-all duration-300
          ${isDragActive
            ? 'border-blue-500 bg-blue-50'
            : disabled
              ? 'border-gray-300 bg-gray-50 cursor-not-allowed'
              : 'border-gray-300 bg-white hover:border-blue-400 hover:bg-blue-50'
          }
        `}
      >
        <input {...getInputProps()} />

        {selectedFiles.length > 0 ? (
          <div className="space-y-2">
            <div className="text-4xl mb-2">📄</div>
            <p className="text-lg font-semibold text-gray-700">
              {selectedFiles.length} file{selectedFiles.length !== 1 ? 's' : ''} selected
            </p>
            <div className="max-h-32 overflow-y-auto">
              {selectedFiles.slice(0, 5).map((file, idx) => (
                <p key={idx} className="text-sm text-gray-500 truncate">
                  {file.name || file.path?.split(/[/\\]/).pop()}
                </p>
              ))}
              {selectedFiles.length > 5 && (
                <p className="text-sm text-gray-400">
                  ... and {selectedFiles.length - 5} more
                </p>
              )}
            </div>
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation()
                onFilesSelected([])
              }}
              className="mt-2 text-sm text-red-500 hover:text-red-700"
            >
              Clear selection
            </button>
          </div>
        ) : (
          <div>
            <div className="text-5xl mb-4">📁</div>
            <p className="text-lg font-medium text-gray-700">{title}</p>
            <p className="text-sm text-gray-500 mt-1">{subtitle}</p>

            {/* Alternative: Use Electron's native file dialog */}
            {window.api && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation()
                  handleBrowse()
                }}
                className="mt-4 px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 transition-colors"
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
