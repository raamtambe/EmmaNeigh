import { useEffect, useState } from 'react'

/**
 * Animated running horse component with status text
 * Uses CSS animation with an ASCII/Unicode horse that "runs"
 */
export default function HorseAnimation({ statusMessage, isRunning = true }) {
  const [frame, setFrame] = useState(0)

  // Horse animation frames (simplified running animation)
  const horseFrames = [
    // Frame 1 - legs extended
    `
      ,--,
   _-'    \`-_
  /   ()    ()\\
 /   /|      |\\\\
    / |  __  | \\
      |_|  |_|
      | |  | |
     _| |  | |_
    `,
    // Frame 2 - legs mid
    `
      ,--,
   _-'    \`-_
  /   ()    ()\\
 /   /|      |\\\\
    / |  __  | \\
      |_|  |_|
       \\|  |/
       _|  |_
    `,
    // Frame 3 - legs together
    `
      ,--,
   _-'    \`-_
  /   ()    ()\\
 /   /|      |\\\\
    / |  __  | \\
      |_|  |_|
       ||  ||
      _||  ||_
    `,
    // Frame 4 - legs extended other way
    `
      ,--,
   _-'    \`-_
  /   ()    ()\\
 /   /|      |\\\\
    / |  __  | \\
      |_|  |_|
      /|  |\\
     _|   |_
    `,
  ]

  useEffect(() => {
    if (!isRunning) return

    const interval = setInterval(() => {
      setFrame((prev) => (prev + 1) % horseFrames.length)
    }, 150)

    return () => clearInterval(interval)
  }, [isRunning])

  return (
    <div className="flex flex-col items-center gap-6 py-8">
      {/* Status message */}
      <div className="text-center">
        <p className="text-xl font-semibold text-blue-600 animate-pulse">
          {statusMessage}
        </p>
      </div>

      {/* Horse container with running effect */}
      <div className="relative w-80 h-48 overflow-hidden bg-gradient-to-b from-blue-50 to-blue-100 rounded-xl border border-blue-200">
        {/* Ground line */}
        <div className="absolute bottom-8 left-0 right-0 h-0.5 bg-amber-600"></div>

        {/* Dust particles */}
        {isRunning && (
          <div className="absolute bottom-8 left-1/4 flex gap-2">
            {[...Array(5)].map((_, i) => (
              <div
                key={i}
                className="w-2 h-2 bg-amber-300 rounded-full opacity-60"
                style={{
                  animation: `dustFloat ${0.5 + i * 0.1}s ease-out infinite`,
                  animationDelay: `${i * 0.1}s`
                }}
              />
            ))}
          </div>
        )}

        {/* Animated Horse SVG */}
        <div
          className={`absolute inset-0 flex items-center justify-center ${isRunning ? 'animate-gallop' : ''}`}
        >
          <svg
            viewBox="0 0 200 150"
            className="w-48 h-36"
            style={{ filter: 'drop-shadow(2px 4px 6px rgba(0,0,0,0.2))' }}
          >
            {/* Horse body - elegant running pose */}
            <g fill="#8B4513" stroke="#5D3A1A" strokeWidth="1">
              {/* Body */}
              <ellipse cx="100" cy="80" rx="45" ry="25" />

              {/* Neck */}
              <path d="M 135 70 Q 155 50 160 35 Q 165 45 155 65 Q 145 75 135 75 Z" />

              {/* Head */}
              <ellipse cx="165" cy="30" rx="18" ry="12" />

              {/* Ear */}
              <path d="M 175 18 L 180 8 L 178 20 Z" />

              {/* Eye */}
              <circle cx="172" cy="28" r="3" fill="#1a1a1a" />

              {/* Mane */}
              <path
                d="M 160 35 Q 150 25 145 40 Q 140 30 135 45 Q 130 35 130 50"
                fill="none"
                stroke="#3D2817"
                strokeWidth="3"
              />

              {/* Tail */}
              <path
                d="M 55 80 Q 35 75 25 90 Q 30 85 40 95 Q 35 100 25 105"
                fill="none"
                stroke="#3D2817"
                strokeWidth="4"
                className={isRunning ? 'animate-pulse' : ''}
              />

              {/* Front legs - animated based on frame */}
              <g transform={`rotate(${frame === 0 ? -20 : frame === 2 ? 20 : 0} 130 95)`}>
                <rect x="125" y="95" width="8" height="35" rx="3" />
                <rect x="125" y="125" width="10" height="8" rx="2" fill="#3D2817" />
              </g>
              <g transform={`rotate(${frame === 2 ? -20 : frame === 0 ? 20 : 0} 115 95)`}>
                <rect x="110" y="95" width="8" height="35" rx="3" />
                <rect x="108" y="125" width="10" height="8" rx="2" fill="#3D2817" />
              </g>

              {/* Back legs - animated based on frame */}
              <g transform={`rotate(${frame === 1 ? -25 : frame === 3 ? 25 : 0} 75 95)`}>
                <rect x="70" y="95" width="8" height="35" rx="3" />
                <rect x="68" y="125" width="10" height="8" rx="2" fill="#3D2817" />
              </g>
              <g transform={`rotate(${frame === 3 ? -25 : frame === 1 ? 25 : 0} 85 95)`}>
                <rect x="80" y="95" width="8" height="35" rx="3" />
                <rect x="78" y="125" width="10" height="8" rx="2" fill="#3D2817" />
              </g>
            </g>
          </svg>
        </div>
      </div>

      {/* Additional CSS for dust animation */}
      <style>{`
        @keyframes dustFloat {
          0% {
            transform: translateX(0) translateY(0) scale(1);
            opacity: 0.6;
          }
          100% {
            transform: translateX(-30px) translateY(-10px) scale(0.3);
            opacity: 0;
          }
        }
      `}</style>
    </div>
  )
}
