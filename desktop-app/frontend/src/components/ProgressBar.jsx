export default function ProgressBar({ percent, showLabel = true }) {
  return (
    <div className="w-full">
      <div className="relative h-4 bg-gray-200 rounded-full overflow-hidden">
        {/* Progress fill */}
        <div
          className="absolute inset-y-0 left-0 bg-gradient-to-r from-blue-500 to-blue-600 transition-all duration-300 ease-out rounded-full"
          style={{ width: `${Math.min(100, Math.max(0, percent))}%` }}
        >
          {/* Animated shine effect */}
          <div className="absolute inset-0 progress-bar-animated" />
        </div>
      </div>

      {showLabel && (
        <div className="flex justify-between mt-2 text-sm text-gray-600">
          <span>Progress</span>
          <span className="font-semibold">{Math.round(percent)}%</span>
        </div>
      )}
    </div>
  )
}
