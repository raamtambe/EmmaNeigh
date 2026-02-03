import { useEffect, useState } from 'react'

/**
 * Clean, minimal horse animation for processing states
 */
export default function HorseAnimation({ statusMessage, isRunning = true }) {
  const [frame, setFrame] = useState(0)

  useEffect(() => {
    if (!isRunning) return
    const interval = setInterval(() => {
      setFrame((prev) => (prev + 1) % 4)
    }, 150)
    return () => clearInterval(interval)
  }, [isRunning])

  return (
    <div className="flex flex-col items-center gap-4 py-6">
      {/* Status message */}
      <p className="text-sm text-slate-600">{statusMessage}</p>

      {/* Horse container */}
      <div className="relative w-64 h-32 overflow-hidden bg-slate-50 rounded-lg">
        {/* Ground line */}
        <div className="absolute bottom-6 left-0 right-0 h-px bg-slate-200" />

        {/* Horse SVG */}
        <div className={`absolute inset-0 flex items-center justify-center ${isRunning ? 'animate-gallop' : ''}`}>
          <svg viewBox="0 0 200 150" className="w-40 h-28">
            <g fill="#475569" stroke="#334155" strokeWidth="1">
              {/* Body */}
              <ellipse cx="100" cy="80" rx="40" ry="22" />

              {/* Neck */}
              <path d="M 130 70 Q 148 52 152 38 Q 156 48 148 65 Q 140 73 130 73 Z" />

              {/* Head */}
              <ellipse cx="156" cy="34" rx="16" ry="10" />

              {/* Ear */}
              <path d="M 165 24 L 169 15 L 168 25 Z" />

              {/* Eye */}
              <circle cx="162" cy="32" r="2" fill="#0f172a" />

              {/* Mane */}
              <path
                d="M 152 38 Q 144 30 140 42 Q 136 34 132 46"
                fill="none"
                stroke="#1e293b"
                strokeWidth="2.5"
              />

              {/* Tail */}
              <path
                d="M 60 80 Q 42 76 34 88 Q 38 84 46 92"
                fill="none"
                stroke="#1e293b"
                strokeWidth="3"
              />

              {/* Front legs */}
              <g transform={`rotate(${frame === 0 ? -15 : frame === 2 ? 15 : 0} 125 92)`}>
                <rect x="120" y="92" width="6" height="30" rx="2" />
                <rect x="118" y="118" width="8" height="6" rx="1" fill="#1e293b" />
              </g>
              <g transform={`rotate(${frame === 2 ? -15 : frame === 0 ? 15 : 0} 112 92)`}>
                <rect x="107" y="92" width="6" height="30" rx="2" />
                <rect x="105" y="118" width="8" height="6" rx="1" fill="#1e293b" />
              </g>

              {/* Back legs */}
              <g transform={`rotate(${frame === 1 ? -18 : frame === 3 ? 18 : 0} 78 92)`}>
                <rect x="73" y="92" width="6" height="30" rx="2" />
                <rect x="71" y="118" width="8" height="6" rx="1" fill="#1e293b" />
              </g>
              <g transform={`rotate(${frame === 3 ? -18 : frame === 1 ? 18 : 0} 88 92)`}>
                <rect x="83" y="92" width="6" height="30" rx="2" />
                <rect x="81" y="118" width="8" height="6" rx="1" fill="#1e293b" />
              </g>
            </g>
          </svg>
        </div>

        {/* Dust dots */}
        {isRunning && (
          <div className="absolute bottom-6 left-1/4 flex gap-1">
            {[...Array(3)].map((_, i) => (
              <div
                key={i}
                className="w-1.5 h-1.5 bg-slate-300 rounded-full"
                style={{
                  animation: `dustFloat 0.6s ease-out infinite`,
                  animationDelay: `${i * 0.1}s`,
                  opacity: 0.5
                }}
              />
            ))}
          </div>
        )}
      </div>

      <style>{`
        @keyframes dustFloat {
          0% {
            transform: translateX(0) translateY(0) scale(1);
            opacity: 0.5;
          }
          100% {
            transform: translateX(-20px) translateY(-8px) scale(0.3);
            opacity: 0;
          }
        }
        @keyframes gallop {
          0%, 100% { transform: translateY(0); }
          50% { transform: translateY(-3px); }
        }
        .animate-gallop {
          animation: gallop 0.3s ease-in-out infinite;
        }
      `}</style>
    </div>
  )
}
