import { useState, useCallback } from 'react'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { 
  convertExcel, 
  downloadResult,
  type ConversionResult 
} from '@/lib/excel-converter'
import { Upload, FileSpreadsheet, Download, CheckCircle, AlertCircle, X, Loader2 } from 'lucide-react'

interface ConversionResultItem {
  file: File
  result: ConversionResult
  mode: 'WASM' | 'JS'
}

interface HyundaiConverterProps {
  wasmStatus: 'loading' | 'ready' | 'fallback'
}

export function HyundaiConverter({ wasmStatus }: HyundaiConverterProps) {
  const [mappingFile, setMappingFile] = useState<File | null>(null)
  const [originFiles, setOriginFiles] = useState<File[]>([])
  const [results, setResults] = useState<ConversionResultItem[]>([])
  const [isProcessing, setIsProcessing] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const handleMappingFileChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setError(null)
    setMappingFile(file)
  }, [])

  const handleOriginFilesChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files
    if (!files) return
    setOriginFiles(Array.from(files))
    setResults([])
  }, [])

  const handleRemoveOriginFile = useCallback((index: number) => {
    setOriginFiles(prev => prev.filter((_, i) => i !== index))
  }, [])

  const handleConvert = useCallback(async () => {
    if (!mappingFile || originFiles.length === 0) return

    setIsProcessing(true)
    setError(null)
    const newResults: ConversionResultItem[] = []

    try {
      for (let i = 0; i < originFiles.length; i++) {
        const file = originFiles[i]
        const { result, mode } = await convertExcel(file, mappingFile)
        newResults.push({ file, result, mode })
        
        // 각 파일 변환 완료 후 바로 다운로드
        await downloadResult(result, file.name)
        
        // 브라우저 연속 다운로드 차단 방지를 위한 지연
        if (i < originFiles.length - 1) {
          await new Promise(resolve => setTimeout(resolve, 500))
        }
      }
      setResults(newResults)
    } catch (err) {
      setError('변환 실패: ' + (err instanceof Error ? err.message : '알 수 없는 오류'))
    } finally {
      setIsProcessing(false)
    }
  }, [mappingFile, originFiles])

  const handleDownload = useCallback(async (file: File, result: ConversionResult) => {
    try {
      await downloadResult(result, file.name)
    } catch (err) {
      setError('다운로드 실패: ' + (err instanceof Error ? err.message : '알 수 없는 오류'))
    }
  }, [])

  const handleDownloadAll = useCallback(async () => {
    for (let i = 0; i < results.length; i++) {
      const { file, result } = results[i]
      await downloadResult(result, file.name)
      
      // 브라우저 연속 다운로드 차단 방지를 위한 지연
      if (i < results.length - 1) {
        await new Promise(resolve => setTimeout(resolve, 500))
      }
    }
  }, [results])

  const canConvert = mappingFile && originFiles.length > 0

  return (
    <div className="space-y-6">
      {/* 설명 */}
      <p className="text-muted-foreground">
        간식서비스 메뉴 엑셀 파일을 시스템 업로드 형식으로 변환합니다
        {wasmStatus === 'ready' && <span className="ml-2 text-xs text-green-600">(WASM 가속)</span>}
        {wasmStatus === 'fallback' && <span className="ml-2 text-xs text-amber-600">(JS 모드)</span>}
      </p>

      {error && (
        <div className="bg-destructive/15 text-destructive px-4 py-3 rounded-md flex items-center gap-2">
          <AlertCircle className="h-4 w-4" />
          <span>{error}</span>
        </div>
      )}

      {/* 매핑 테이블 */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <FileSpreadsheet className="h-5 w-5" />
            매핑 테이블
          </CardTitle>
          <CardDescription>
            원본 사업장명과 시스템 사업장명 매핑 파일을 업로드하세요
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="grid w-full items-center gap-3">
            <Label htmlFor="mapping-file">매핑 테이블 파일 (.xlsx)</Label>
            <Input
              id="mapping-file"
              type="file"
              accept=".xlsx,.xls"
              onChange={handleMappingFileChange}
            />
            {mappingFile && (
              <div className="flex items-center gap-2 text-sm text-green-600">
                <CheckCircle className="h-4 w-4" />
                <span>{mappingFile.name}</span>
              </div>
            )}
          </div>
        </CardContent>
      </Card>

      {/* 원본 파일 */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Upload className="h-5 w-5" />
            원본 파일
          </CardTitle>
          <CardDescription>
            변환할 원본 엑셀 파일을 업로드하세요 (여러 개 선택 가능)
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="grid w-full items-center gap-3">
            <Label htmlFor="origin-files">원본 파일 (.xlsx, 여러 개 선택 가능)</Label>
            <Input
              id="origin-files"
              type="file"
              accept=".xlsx,.xls"
              multiple
              onChange={handleOriginFilesChange}
            />
            {originFiles.length > 0 && (
              <div className="space-y-2">
                {originFiles.map((file, idx) => (
                  <div key={idx} className="flex items-center justify-between gap-2 text-sm bg-muted px-3 py-2 rounded-md">
                    <span className="truncate">{file.name}</span>
                    <button
                      onClick={() => handleRemoveOriginFile(idx)}
                      className="text-muted-foreground hover:text-foreground"
                    >
                      <X className="h-4 w-4" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </CardContent>
      </Card>

      {/* 변환 버튼 */}
      <div className="flex justify-center">
        <Button
          size="lg"
          onClick={handleConvert}
          disabled={!canConvert || isProcessing}
        >
          {isProcessing ? (
            <>
              <Loader2 className="h-4 w-4 mr-2 animate-spin" />
              변환 중...
            </>
          ) : (
            '변환하기'
          )}
        </Button>
      </div>

      {/* 결과 영역 */}
      {results.length > 0 && (
        <div className="space-y-6 pt-6 border-t">
          <div className="flex items-center justify-between">
            <h3 className="text-lg font-semibold">변환 결과</h3>
            <Button onClick={handleDownloadAll} variant="outline" size="sm">
              <Download className="h-4 w-4 mr-2" />
              전체 다운로드
            </Button>
          </div>

          {results.map(({ file, result, mode }, idx) => (
            <Card key={idx}>
              <CardHeader>
                <CardTitle className="text-lg flex items-center justify-between">
                  <span>{file.name}</span>
                  <span className="text-xs font-normal text-muted-foreground bg-muted px-2 py-1 rounded">
                    {mode}
                  </span>
                </CardTitle>
                <CardDescription>
                  총 {result.data.length}개 데이터 추출
                  {result.mappingFailures.length > 0 && (
                    <span className="text-destructive ml-2">
                      (매핑 실패: {result.mappingFailures.length}개 매장)
                    </span>
                  )}
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                {/* 검증 테이블 */}
                <div className="rounded-md border overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-muted">
                      <tr>
                        <th className="px-4 py-2 text-left font-medium">일자</th>
                        <th className="px-4 py-2 text-left font-medium">요일</th>
                        <th className="px-4 py-2 text-right font-medium">추출 Box</th>
                        <th className="px-4 py-2 text-right font-medium">원본 Box</th>
                        <th className="px-4 py-2 text-left font-medium">결과</th>
                      </tr>
                    </thead>
                    <tbody>
                      {result.validation.map((row, vIdx) => (
                        <tr key={vIdx} className="border-t">
                          <td className="px-4 py-2">{row.date}</td>
                          <td className="px-4 py-2">{row.dayName}</td>
                          <td className="px-4 py-2 text-right">{row.extractedBox}</td>
                          <td className="px-4 py-2 text-right">{row.originalBox}</td>
                          <td className="px-4 py-2">
                            <span className={row.result === '일치' ? 'text-green-600' : 'text-destructive'}>
                              {row.result}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* 매핑 실패 목록 */}
                {result.mappingFailures.length > 0 && (
                  <div className="bg-destructive/10 p-4 rounded-md">
                    <h4 className="font-medium text-destructive mb-2">매핑 실패 매장</h4>
                    <ul className="text-sm space-y-1">
                      {result.mappingFailures.map((name, fIdx) => (
                        <li key={fIdx}>• {name}</li>
                      ))}
                    </ul>
                  </div>
                )}

                <div className="flex justify-end">
                  <Button onClick={() => handleDownload(file, result)} size="sm">
                    <Download className="h-4 w-4 mr-2" />
                    다운로드
                  </Button>
                </div>
              </CardContent>
            </Card>
          ))}
        </div>
      )}
    </div>
  )
}
