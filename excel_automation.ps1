
# Excel 자동화 스크립트 (개선된 버전)
$excelFilePath = "C:\\Users\\nepes\\Downloads\\DynamicsExport_638872039208168819.xlsx"
$vbaFilePath = "C:\\Users\\nepes\\OneDrive - 네패스\\바탕 화면\\PROJECT_GIT\\Electron_Test\\my-electron-app\\Tax_Invoice_app\\temp_macro.vba"

Write-Host "🚀 Excel 자동화 스크립트 시작"
Write-Host "파일 경로: $excelFilePath"

# 기존 Excel 프로세스 확인
$existingExcel = $null
try {
    $existingExcel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    Write-Host "✅ 기존 Excel 애플리케이션 발견"
} catch {
    Write-Host "📂 새로운 Excel 애플리케이션 시작"
    $existingExcel = New-Object -ComObject Excel.Application
}

$excel = $existingExcel
$excel.Visible = $true
$excel.DisplayAlerts = $false

try {
    Write-Host "📂 엑셀 파일 열기 중..."
    
    # 이미 열린 파일인지 확인
    $workbook = $null
    $fileName = [System.IO.Path]::GetFileName($excelFilePath)
    
    foreach($wb in $excel.Workbooks) {
        if ($wb.Name -eq $fileName) {
            $workbook = $wb
            Write-Host "✅ 이미 열린 파일 발견: $fileName"
            break
        }
    }
    
    # 파일이 열려있지 않으면 새로 열기
    if ($workbook -eq $null) {
        $workbook = $excel.Workbooks.Open($excelFilePath)
        Write-Host "✅ 엑셀 파일 새로 열기 완료"
    }
    
    # 파일 로딩 완료 대기
    Start-Sleep -Seconds 5
    
    Write-Host "📝 VBA 매크로 코드 읽기 중..."
    
    # VBA 코드 읽기
    $vbaCode = Get-Content -Path $vbaFilePath -Raw
    
    Write-Host "🔧 VBA 매크로 모듈 추가 중..."
    
    # 기존 모듈이 있는지 확인하고 제거
    $vbaProject = $workbook.VBProject
    $existingModule = $null
    
    foreach($component in $vbaProject.VBComponents) {
        if ($component.Name -eq "AutoGeneratedModule") {
            $existingModule = $component
            break
        }
    }
    
    if ($existingModule -ne $null) {
        $vbaProject.VBComponents.Remove($existingModule)
        Write-Host "🗑️ 기존 매크로 모듈 제거"
    }
    
    # VBA 프로젝트에 새 모듈 추가
    $vbaModule = $vbaProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
    $vbaModule.Name = "AutoGeneratedModule"
    
    # VBA 코드 추가
    $vbaModule.CodeModule.AddFromString($vbaCode)
    
    Write-Host "✅ VBA 매크로 모듈 추가 완료"
    Write-Host "▶️ 매크로 실행 중..."
    
    # 매크로 실행
    $excel.Run("AutoGeneratedModule.GroupBy_I_Z_And_Process")
    
    Write-Host "🎉 매크로 실행 완료!"
    
    # 파일 저장
    $workbook.Save()
    Write-Host "💾 엑셀 파일 저장 완료"
    
    Write-Host "🎯 모든 작업이 성공적으로 완료되었습니다!"
    
} catch {
    Write-Host "❌ 오류 발생: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "❌ 상세 오류: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    throw
} finally {
    # Excel 애플리케이션 종료하지 않음 (사용자가 결과 확인할 수 있도록)
    Write-Host "📊 Excel 애플리케이션을 열어둡니다. 결과를 확인하세요."
    Write-Host "🏁 스크립트 실행 완료"
}
