import { useNavigate } from 'react-router-dom'

// Horse icon SVG component
function HorseIcon({ className }) {
  return (
    <svg
      className={className}
      viewBox="0 0 100 100"
      fill="currentColor"
    >
      <path d="M85 25c-5-8-15-10-25-8-3-5-8-7-15-7-10 0-18 5-22 12-8 2-14 8-16 16-3 12 2 25 12 32v15c0 3 2 5 5 5h8c3 0 5-2 5-5v-10h10v10c0 3 2 5 5 5h8c3 0 5-2 5-5V70c8-5 14-14 15-25 8 0 12-10 5-20zM35 45c-3 0-5-2-5-5s2-5 5-5 5 2 5 5-2 5-5 5z"/>
    </svg>
  )
}

function FeatureCard({ title, description, icon, onClick, disabled, comingSoon }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`
        relative p-8 rounded-2xl text-left transition-all duration-300
        ${disabled
          ? 'bg-gray-100 cursor-not-allowed opacity-60'
          : 'bg-white hover:bg-blue-50 hover:shadow-xl hover:-translate-y-1 cursor-pointer shadow-lg'
        }
        border border-gray-200
      `}
    >
      {comingSoon && (
        <span className="absolute top-4 right-4 bg-amber-100 text-amber-700 text-xs font-semibold px-2 py-1 rounded-full">
          Coming Soon
        </span>
      )}
      <div className="text-5xl mb-4">{icon}</div>
      <h3 className="text-xl font-bold text-emma-navy mb-2">{title}</h3>
      <p className="text-gray-600 text-sm">{description}</p>
    </button>
  )
}

export default function MainMenu() {
  const navigate = useNavigate()

  return (
    <div className="min-h-screen flex flex-col items-center justify-center p-8">
      {/* Header */}
      <div className="text-center mb-12 animate-fade-in">
        <div className="flex items-center justify-center gap-4 mb-4">
          <HorseIcon className="w-16 h-16 text-emma-navy" />
          <h1 className="text-5xl font-bold text-emma-navy">EmmaNeigh</h1>
        </div>
        <p className="text-xl text-gray-600">Signature Packet Automation</p>
        <p className="text-sm text-gray-500 mt-2">For M&A Transactions and Financing Deals</p>
      </div>

      {/* Feature Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 max-w-5xl w-full">
        <FeatureCard
          icon="📝"
          title="Create Signature Packets"
          description="Extract signature pages from multiple PDFs and organize them by signer. Creates individual packets ready for DocuSign."
          onClick={() => navigate('/signature-packets')}
        />

        <FeatureCard
          icon="📋"
          title="Create Execution Version"
          description="Merge signed pages back into the original document. Automatically unlocks DocuSign PDFs and creates the final execution version."
          onClick={() => navigate('/execution-version')}
        />

        <FeatureCard
          icon="📧"
          title="DocuSign Integration"
          description="Send signature packets directly to signers via DocuSign. Track signature status and receive completed documents."
          disabled
          comingSoon
        />
      </div>

      {/* Footer */}
      <div className="mt-12 text-center text-sm text-gray-500">
        <p>All processing is done locally on your machine.</p>
        <p className="mt-1">No data is ever sent to the cloud.</p>
      </div>
    </div>
  )
}
