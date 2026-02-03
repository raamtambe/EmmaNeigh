import { useNavigate } from 'react-router-dom'

// Clean, professional horse icon
function HorseIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="currentColor">
      <path d="M21.5 9.5c-.3-1.1-1-2-1.9-2.6.2-.9.1-1.8-.3-2.6-.4-.8-1.1-1.4-2-1.7-.2-.1-.4-.1-.6 0-.2.1-.3.2-.4.4l-.8 2c-.3.1-.6.2-.9.4l-2-.8c-.2-.1-.4-.1-.6 0-.2.1-.3.2-.4.4-.6 1.4-.2 3 .9 4l-.3.5c-.9.2-1.7.7-2.3 1.4l-6.2 6.2c-.4.4-.4 1 0 1.4.2.2.5.3.7.3s.5-.1.7-.3l5.8-5.8c.1.3.2.6.2.9l-1.5 4.5c-.1.4 0 .8.3 1.1.2.2.5.3.8.3.1 0 .3 0 .4-.1l3-1c.2-.1.4-.2.5-.4l2.6-3.9c.4-.1.8-.2 1.1-.4l1.8.9c.1.1.3.1.4.1.3 0 .5-.1.7-.3.3-.3.4-.7.2-1.1l-.9-1.8c.2-.4.3-.8.4-1.3h2c.6 0 1-.4 1-1s-.4-1-1-1zM19 10.5c0 .3 0 .5-.1.8l-.1.4.8 1.6-1.2-.6-.3.2c-.4.3-.8.4-1.3.5l-.4.1-2.5 3.8-2 .7 1.2-3.6.1-.4c0-.5-.1-1-.3-1.4l-.2-.4.3-.3c.5-.5 1.1-.8 1.8-.9l.4-.1.7-1.3-.1-.4c-.3-.6-.4-1.2-.2-1.8l1.4.6.4-.2c.4-.2.9-.4 1.3-.4l.4-.1.6-1.4c.3.2.5.5.6.8.2.5.2 1.1 0 1.6l-.2.4.3.3c.5.5.8 1.1.9 1.8z"/>
    </svg>
  )
}

// Document icon
function DocumentIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
    </svg>
  )
}

// Merge icon
function MergeIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path strokeLinecap="round" strokeLinejoin="round" d="M7.5 21L3 16.5m0 0L7.5 12M3 16.5h13.5m0-13.5L21 7.5m0 0L16.5 12M21 7.5H7.5" />
    </svg>
  )
}

// Send icon
function SendIcon({ className }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path strokeLinecap="round" strokeLinejoin="round" d="M6 12L3.269 3.126A59.768 59.768 0 0121.485 12 59.77 59.77 0 013.27 20.876L5.999 12zm0 0h7.5" />
    </svg>
  )
}

function FeatureCard({ title, description, icon: Icon, onClick, disabled, comingSoon }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`
        relative p-6 rounded-lg text-left transition-all duration-200
        ${disabled
          ? 'bg-slate-50 cursor-not-allowed'
          : 'bg-white hover:bg-slate-50 cursor-pointer border border-slate-200 hover:border-slate-300 hover:shadow-sm'
        }
      `}
    >
      {comingSoon && (
        <span className="absolute top-3 right-3 text-xs font-medium text-slate-400 bg-slate-100 px-2 py-0.5 rounded">
          Soon
        </span>
      )}
      <div className={`w-10 h-10 rounded-lg flex items-center justify-center mb-4 ${disabled ? 'bg-slate-100' : 'bg-slate-900'}`}>
        <Icon className={`w-5 h-5 ${disabled ? 'text-slate-400' : 'text-white'}`} />
      </div>
      <h3 className={`text-base font-semibold mb-1 ${disabled ? 'text-slate-400' : 'text-slate-900'}`}>{title}</h3>
      <p className={`text-sm leading-relaxed ${disabled ? 'text-slate-400' : 'text-slate-500'}`}>{description}</p>
    </button>
  )
}

export default function MainMenu() {
  const navigate = useNavigate()

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      {/* Header */}
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-3xl mx-auto px-6 py-6">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-slate-900 rounded-lg flex items-center justify-center">
              <HorseIcon className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-semibold text-slate-900">EmmaNeigh</h1>
              <p className="text-sm text-slate-500">Signature Packet Automation</p>
            </div>
          </div>
        </div>
      </header>

      {/* Main content */}
      <main className="flex-1 max-w-3xl mx-auto w-full px-6 py-8">
        <div className="mb-6">
          <h2 className="text-sm font-medium text-slate-500 uppercase tracking-wide">Tools</h2>
        </div>

        {/* Feature Cards */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <FeatureCard
            icon={DocumentIcon}
            title="Create Signature Packets"
            description="Extract signature pages from PDFs and organize by signer for DocuSign."
            onClick={() => navigate('/signature-packets')}
          />

          <FeatureCard
            icon={MergeIcon}
            title="Create Execution Version"
            description="Merge signed pages back into the original document."
            onClick={() => navigate('/execution-version')}
          />

          <FeatureCard
            icon={SendIcon}
            title="DocuSign Integration"
            description="Send packets directly to signers via DocuSign."
            disabled
            comingSoon
          />
        </div>
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-200 bg-white">
        <div className="max-w-3xl mx-auto px-6 py-4">
          <p className="text-xs text-slate-400">
            All processing happens locally on your machine. No data is sent to external servers.
          </p>
        </div>
      </footer>
    </div>
  )
}
