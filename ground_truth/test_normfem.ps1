# Тестовый PowerShell скрипт для проверки API NormFEM
# Запуск (32-bit): C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File test_normfem.ps1

Write-Host "======================================"
Write-Host "ТЕСТ API NormFEM (PowerShell + данные)"
Write-Host "======================================"

try {
    Write-Host ""
    Write-Host "[INFO] Создание COM объекта ncfem.main..."
    $nfApi = New-Object -ComObject "ncfem.main"
    Write-Host "[OK] COM объект создан"
}
catch {
    Write-Host "[FAIL] Ошибка создания COM: $_"
    exit 1
}

try {
    Write-Host ""
    Write-Host "[INFO] SetPath..."
    $nfApi.SetPath("C:\Program Files (x86)\NormCAD", "C:\Temp", "TestProject", $null)
    Write-Host "[OK] SetPath выполнен"
}
catch {
    Write-Host "[FAIL] SetPath: $_"
}

# ========== ПЕРЕДАЧА ДАННЫХ ==========
Write-Host ""
Write-Host "[INFO] Передача данных через SetArr..."

# Создание типизированных массивов String
try {
    [string[]]$arrM00 = @("Steel; 2.0e5; 0.3")
    $nfApi.SetArr("m00", $arrM00)
    Write-Host "[OK] m00 (материалы): $($arrM00.Count) шт"
}
catch {
    Write-Host "[FAIL] m00: $_"
}

try {
    [string[]]$arrG01 = @("1; 0; 0; 0", "2; 3000; 0; 0")
    $nfApi.SetArr("g01", $arrG01)
    Write-Host "[OK] g01 (узлы): $($arrG01.Count) шт"
}
catch {
    Write-Host "[FAIL] g01: $_"
}

try {
    [string[]]$arrG02 = @("1; 1; 2; 1; 40000; 1.33e8; 1.33e8; 2.67e8")
    $nfApi.SetArr("g02", $arrG02)
    Write-Host "[OK] g02 (элементы): $($arrG02.Count) шт"
}
catch {
    Write-Host "[FAIL] g02: $_"
}

try {
    [string[]]$arrG03 = @("1; 1; 1; 1; 0; 0; 0", "2; 0; 1; 1; 0; 0; 0")
    $nfApi.SetArr("g03", $arrG03)
    Write-Host "[OK] g03 (связи): $($arrG03.Count) шт"
}
catch {
    Write-Host "[FAIL] g03: $_"
}

try {
    [string[]]$arrL00 = @("1; 1; 0; 0; -1.5; 0; 0; 0")
    $nfApi.SetArr("l00", $arrL00)
    Write-Host "[OK] l00 (нагрузки): $($arrL00.Count) шт"
}
catch {
    Write-Host "[FAIL] l00: $_"
}

# Дополнительные таблицы которые могут быть нужны
try {
    # s00 - сечения (если требуется)
    [string[]]$arrS00 = @("1; Rectangle; 200; 200")
    $nfApi.SetArr("s00", $arrS00)
    Write-Host "[OK] s00 (сечения): $($arrS00.Count) шт"
}
catch {
    Write-Host "[INFO] s00: не требуется или $_"
}

try {
    # Свойство Steps
    $nfApi.Steps = 5
    Write-Host "[OK] Steps = 5"
}
catch {
    Write-Host "[INFO] Steps: $_"
}

try {
    # Проверка Mode3D
    $mode = $nfApi.Mode3D
    Write-Host "[INFO] Mode3D = $mode"
}
catch {
    Write-Host "[INFO] Mode3D: $_"
}

# ========== РАСЧЁТ ==========
Write-Host ""
Write-Host "[INFO] Вызов Calc..."
try {
    $nfApi.Calc()
    Write-Host "[OK] Calc выполнен"
}
catch {
    Write-Host "[FAIL] Calc: $_"
}

# Result
Write-Host ""
Write-Host "[INFO] Проверка Result..."
try {
    $result = $nfApi.Result
    Write-Host "[OK] Result = $result"
}
catch {
    Write-Host "[FAIL] Result: $_"
}

# ========== ПОЛУЧЕНИЕ РЕЗУЛЬТАТОВ ==========
Write-Host ""
Write-Host "[INFO] Получение результатов..."

try {
    $ArrZ = $null
    $nfApi.GetArrZ([ref]$ArrZ)
    Write-Host "[OK] ArrZ получен"
}
catch {
    Write-Host "[FAIL] GetArrZ: $_"
}

try {
    $ArrNM = $null
    $nfApi.GetArrNM([ref]$ArrNM)
    Write-Host "[OK] ArrNM получен"
}
catch {
    Write-Host "[FAIL] GetArrNM: $_"
}

try {
    $ArrQ = $null
    $nfApi.GetArrQ([ref]$ArrQ)
    Write-Host "[OK] ArrQ получен"
}
catch {
    Write-Host "[FAIL] GetArrQ: $_"
}

# Освобождение
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($nfApi) | Out-Null

Write-Host ""
Write-Host "======================================"
Write-Host "ТЕСТ ЗАВЕРШЁН"
Write-Host "======================================"
