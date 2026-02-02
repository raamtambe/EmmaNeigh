import { Routes, Route } from 'react-router-dom'
import MainMenu from './pages/MainMenu'
import SignaturePackets from './pages/SignaturePackets'
import ExecutionVersion from './pages/ExecutionVersion'

function App() {
  return (
    <div className="min-h-screen bg-emma-cream">
      <Routes>
        <Route path="/" element={<MainMenu />} />
        <Route path="/signature-packets" element={<SignaturePackets />} />
        <Route path="/execution-version" element={<ExecutionVersion />} />
      </Routes>
    </div>
  )
}

export default App
