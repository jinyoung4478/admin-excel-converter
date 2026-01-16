import { useEffect, useState } from 'react'
import { BrowserRouter, Routes, Route, useNavigate, useLocation } from 'react-router-dom'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { HyundaiConverter } from '@/features/hyundai/HyundaiConverter'
import { initializeWasm } from '@/lib/excel-converter'
import { CheckCircle, Loader2 } from 'lucide-react'

// 탭과 라우트 매핑
const TAB_ROUTES = {
  hyundai: '/hyundai',
} as const

type TabKey = keyof typeof TAB_ROUTES

function AppContent() {
  const navigate = useNavigate()
  const location = useLocation()
  const [wasmStatus, setWasmStatus] = useState<'loading' | 'ready' | 'fallback'>('loading')

  // WASM 초기화
  useEffect(() => {
    initializeWasm().then(success => {
      setWasmStatus(success ? 'ready' : 'fallback')
    })
  }, [])

  // 현재 경로에서 활성 탭 결정
  const getActiveTab = (): TabKey | '' => {
    if (location.pathname === '/hyundai') return 'hyundai'
    return ''
  }

  const activeTab = getActiveTab()

  // 탭 변경 시 라우트 변경
  const handleTabChange = (value: string) => {
    if (value in TAB_ROUTES) {
      navigate(TAB_ROUTES[value as TabKey])
    }
  }

  // 기본 경로('/')일 때 hyundai로 리다이렉트
  useEffect(() => {
    if (location.pathname === '/') {
      navigate('/hyundai', { replace: true })
    }
  }, [location.pathname, navigate])

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b">
        <div className="container mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <h1 className="text-2xl font-bold">엑셀 변환기</h1>
            <div className="text-xs text-muted-foreground">
              {wasmStatus === 'loading' && (
                <span className="flex items-center gap-1">
                  <Loader2 className="h-3 w-3 animate-spin" />
                  WASM 로딩...
                </span>
              )}
              {wasmStatus === 'ready' && (
                <span className="flex items-center gap-1 text-green-600">
                  <CheckCircle className="h-3 w-3" />
                  WASM 활성화
                </span>
              )}
              {wasmStatus === 'fallback' && (
                <span className="flex items-center gap-1 text-yellow-600">
                  JS 모드
                </span>
              )}
            </div>
          </div>
        </div>
      </header>

      <main className="container mx-auto px-4 py-6">
        <Tabs value={activeTab} onValueChange={handleTabChange}>
          <TabsList>
            <TabsTrigger value="hyundai">현대차</TabsTrigger>
            {/* 추후 다른 변환 기능 추가 시 여기에 TabsTrigger 추가 */}
          </TabsList>

          <TabsContent value="hyundai" className="mt-6">
            <HyundaiConverter wasmStatus={wasmStatus} />
          </TabsContent>

          {/* 추후 다른 변환 기능 추가 시 여기에 TabsContent 추가 */}
        </Tabs>

        {/* 탭이 선택되지 않은 경우 (라우터로 인해 거의 발생하지 않음) */}
        {!activeTab && location.pathname === '/' && (
          <div className="text-center py-12 text-muted-foreground">
            변환할 서비스를 선택하세요
          </div>
        )}
      </main>
    </div>
  )
}

function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/*" element={<AppContent />} />
      </Routes>
    </BrowserRouter>
  )
}

export default App
