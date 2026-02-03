export default function ProgressBar({ percent, showLabel = true }) {
  return (
    <div className="w-full">
      <div className="relative h-1.5 bg-slate-200 rounded-full overflow-hidden">
        <div
          className="absolute inset-y-0 left-0 bg-slate-600 transition-all duration-300 ease-out rounded-full"
          style={{ width: `${Math.min(100, Math.max(0, percent))}%` }}
        />
      </div>

      {showLabel && (
        <div className="flex justify-end mt-2">
          <span className="text-xs text-slate-500">{Math.round(percent)}%</span>
        </div>
      )}
    </div>
  )
}
