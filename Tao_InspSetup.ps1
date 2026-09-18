# Tao_InspSetup.ps1 v4 - Đồng bộ inspection setup Plant 1000 theo quy tắc TTK
#   TAO       : tạo loại kiểm còn thiếu theo quy tắc
#   SUA_DM1   : (-SuaDM1) TTK10 chưa có dynamic rule DM1 -> DM1
#   TAT       : (-DonDep) loại kiểm active nằm ngoài quy tắc nhóm & ngoài $GiuLai -> inactive
#   KICH_HOAT : (-DonDep) loại kiểm trong quy tắc nhưng đang inactive -> active
# v3: credential env/DPAPI; không có việc thì thoát, không tạo log; chỉ cài ImportExcel khi cần.
# v4: thêm -DonDep (API không cho xóa -> tắt bằng ProdInspTypeSettingIsActive=false).
# Mặc định DRY-RUN. -ThucHien để ghi. -GioiHan N chạy thử N dòng. -KichThuocLo 1 = gửi lẻ từng dòng.
param(
  [switch]$ThucHien,
  [switch]$SuaDM1,
  [switch]$DonDep,
  [int]$GioiHan = 0,
  [int]$KichThuocLo = 50
)
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'   # tăng tốc Invoke-WebRequest trên PS 5.1
$BaseUrl  = 'https://my412079-api.s4hana.cloud.sap'
$Svc      = "$BaseUrl/sap/opu/odata4/sap/api_product/srvd_a2x/sap/product/0003"
$Plant    = '1000'
$CredFile = Join-Path $PSScriptRoot 'sap_cred.xml'   # <-- chỉ dùng khi chạy trên máy

$TP = @('TTK04','TTK10','TTK89','TTK89KPH','TTK89TV')
$QuyTac = [ordered]@{
  'M020'=@{Tu=20000000;Den=20999999;Loai=@('TTK01','TTK89')}
  'M021'=@{Tu=21000000;Den=21999999;Loai=@('TTK01','TTK89')}
  'M041'=@{Tu=41000000;Den=41999999;Loai=@('TTK04','TTK89')}
  'M042'=@{Tu=42000000;Den=42999999;Loai=@('TTK01','TTK04','TTK89')}
  'M044'=@{Tu=44000000;Den=44999999;Loai=@('TTK04','TTK89')}
  'M050'=@{Tu=50000000;Den=50999999;Loai=$TP}
  'M051'=@{Tu=51000000;Den=51999999;Loai=$TP}
  'M052'=@{Tu=52000000;Den=52999999;Loai=$TP}
}
# Loại kiểm được giữ nguyên dù nằm ngoài quy tắc (không bị -DonDep tắt). Để trống = tắt hết loại ngoài quy tắc.
$GiuLai = @()

function Get-Body($ma, $loai, $nhom) {
  $b = [ordered]@{
    Product=$ma; Plant=$Plant; InspectionLotType=$loai; ProdInspTypeSettingIsActive=$true
    InspLotIsTaskListRequired=$true; InspLotHasAutomSpecAssgmt=$true; InspLotHasCharc=$true
    HasPostToInspectionStock=$false; InspLotIsAutomUsgeDcsnPossible=$true; InspLotSkipIsAllowed=$true
    InspQualityScoreProcedure='06'; InspTypeIsPrfrd=$false
  }
  switch ($loai) {
    'TTK01' { $b.InspTypeIsPrfrd=$true; $b.InspLotSummaryControl='X' }
    'TTK04' { $b.InspTypeIsPrfrd=$true; if ($nhom -eq 'M042') { $b.InspLotSummaryControl='X' } }
    'TTK10' { $b.InspTypeIsPrfrd=$true; $b.InspLotSummaryControl='3'; $b.InspLotDynamicRule='DM1' }
  }
  return ($b | ConvertTo-Json -Compress)
}

# ---- Auth ----
if ($env:SAP_USER -and $env:SAP_PASS) {            # chạy trên GitHub Actions
  $pair = "{0}:{1}" -f $env:SAP_USER, $env:SAP_PASS
} else {                                           # chạy trên máy (DPAPI)
  $cred = Import-Clixml $CredFile
  $pair = "{0}:{1}" -f $cred.UserName, $cred.GetNetworkCredential().Password
}
$auth = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
$hdrJ = @{ Authorization=$auth; Accept='application/json' }

function Get-All($path) {
  $all = New-Object System.Collections.Generic.List[object]; $skip = 0; $top = 5000
  do {
    $sep = if ($path -match '\?') { '&' } else { '?' }
    $url = "$Svc/$path$sep`$top=$top&`$skip=$skip"; Write-Host "GET $url"
    $r = Invoke-RestMethod -Uri $url -Headers $hdrJ
    foreach ($x in $r.value) { $all.Add($x) }; $skip += $top
  } while ($r.value.Count -eq $top)
  return $all
}
function NhomCua($p) {
  $n=0L; if (-not [long]::TryParse($p.TrimStart('0'),[ref]$n)) { return $null }
  foreach ($k in $QuyTac.Keys) { if ($n -ge $QuyTac[$k].Tu -and $n -le $QuyTac[$k].Den) { return $k } }
  return $null
}
function Get-Csrf {
  $r = Invoke-WebRequest -Uri "$Svc/" -Headers @{Authorization=$auth;'x-csrf-token'='Fetch';Accept='application/json'} `
         -SessionVariable s -UseBasicParsing
  $script:ss = $s; $script:csrf = $r.Headers['x-csrf-token']
}

# ---- Gửi 1 lô qua $batch, trả về mảng kết quả theo đúng thứ tự ----
function Send-Batch($lo) {
  $bd = 'batch_' + [guid]::NewGuid().ToString('N')
  $sb = New-Object Text.StringBuilder
  $n = 0
  foreach ($v in $lo) {
    $n++
    if ($v.HanhDong -eq 'TAO') {
      $method = 'POST'; $rel = "ProductPlant(Product='$($v.Ma)',Plant='$Plant')/_ProductPlantInspTypeSetting"
      $body = Get-Body $v.Ma $v.Loai $v.Nhom; $extra = ''
    } else {
      $method = 'PATCH'; $rel = "ProductPlantInspTypeSetting(Product='$($v.Ma)',Plant='$Plant',InspectionLotType='$($v.Loai)')"
      $extra = "If-Match: $($v.Etag)`r`n"
      $body = switch ($v.HanhDong) {
        'SUA_DM1'   { '{"InspLotDynamicRule":"DM1"}' }
        'TAT'       { '{"ProdInspTypeSettingIsActive":false}' }
        'KICH_HOAT' { '{"ProdInspTypeSettingIsActive":true}' }
      }
    }
    [void]$sb.Append("--$bd`r`nContent-Type: application/http`r`nContent-Transfer-Encoding: binary`r`nContent-ID: $n`r`n`r`n")
    [void]$sb.Append("$method $rel HTTP/1.1`r`nContent-Type: application/json`r`nAccept: application/json`r`n$extra`r`n$body`r`n")
  }
  [void]$sb.Append("--$bd--`r`n")

  for ($lan=1; $lan -le 2; $lan++) {
    try {
      $res = Invoke-WebRequest -Uri "$Svc/`$batch" -Method Post -WebSession $script:ss `
               -Body ([Text.Encoding]::UTF8.GetBytes($sb.ToString())) `
               -Headers @{ Authorization=$auth; 'x-csrf-token'=$script:csrf; Accept='multipart/mixed' } `
               -ContentType "multipart/mixed; boundary=$bd" -UseBasicParsing
      break
    } catch {
      $code = [int]$_.Exception.Response.StatusCode
      if ($code -eq 403 -and $lan -eq 1) { Get-Csrf; continue }
      $msg = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
      return @($lo | ForEach-Object { @{ OK=$false; Code=$code; Msg="Lỗi cả lô: $msg" } })
    }
  }

  $txt = if ($res.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($res.Content) } else { $res.Content }
  $rbd = ([regex]::Match($res.Headers['Content-Type'], 'boundary=([^;\s]+)')).Groups[1].Value.Trim('"')
  $parts = $txt -split [regex]::Escape("--$rbd") | Where-Object { $_ -match 'HTTP/1\.1 \d{3}' }
  $kq = foreach ($p in $parts) {
    $c = [int]([regex]::Match($p, 'HTTP/1\.1 (\d{3})').Groups[1].Value)
    $m = ''
    if ($c -ge 400) {
      $j = ([regex]::Match($p, '\{"error".*\}', 'Singleline')).Value
      try { $m = ($j | ConvertFrom-Json).error.message } catch { $m = $j }
    }
    @{ OK=($c -lt 400); Code=$c; Msg=$m }
  }
  if (@($kq).Count -ne $lo.Count) {
    return @($lo | ForEach-Object { @{ OK=$false; Code='?'; Msg="Response batch không khớp số dòng ($(@($kq).Count)/$($lo.Count)) - chạy lại để đối chiếu" } })
  }
  return @($kq)
}

# ---- 1. Đối chiếu ----
$maPlant = @(Get-All "ProductPlant?`$filter=Plant eq '$Plant'&`$select=Product,Plant,IsMarkedForDeletion" |
             Where-Object { -not $_.IsMarkedForDeletion })
$setup = Get-All "ProductPlantInspTypeSetting?`$filter=Plant eq '$Plant'&`$select=Product,InspectionLotType,ProdInspTypeSettingIsActive,InspLotDynamicRule"
$idx = @{}; foreach ($s in $setup) { $k=$s.Product; if (-not $idx[$k]) { $idx[$k]=@{} }; $idx[$k][$s.InspectionLotType]=$s }

$viec = New-Object System.Collections.Generic.List[object]
foreach ($m in $maPlant) {
  $nhom = NhomCua $m.Product; if (-not $nhom) { continue }
  $hc = if ($idx[$m.Product]) { $idx[$m.Product] } else { @{} }
  $quyTacNhom = $QuyTac[$nhom].Loai
  foreach ($l in $quyTacNhom) {
    if (-not $hc.ContainsKey($l)) {
      $viec.Add([pscustomobject]@{ Ma=$m.Product; Nhom=$nhom; Loai=$l; HanhDong='TAO'; Etag='' })
      continue
    }
    $s = $hc[$l]
    if ($DonDep -and -not $s.ProdInspTypeSettingIsActive) {
      $viec.Add([pscustomobject]@{ Ma=$m.Product; Nhom=$nhom; Loai=$l; HanhDong='KICH_HOAT'; Etag=$s.'@odata.etag' })
    }
    if ($SuaDM1 -and $l -eq 'TTK10' -and $s.InspLotDynamicRule -ne 'DM1') {
      if ($DonDep -and -not $s.ProdInspTypeSettingIsActive) { continue }   # tránh 2 PATCH cùng dòng/etag; lần chạy sau sẽ sửa DM1
      $viec.Add([pscustomobject]@{ Ma=$m.Product; Nhom=$nhom; Loai=$l; HanhDong='SUA_DM1'; Etag=$s.'@odata.etag' })
    }
  }
  if ($DonDep) {
    foreach ($l in $hc.Keys) {
      $s = $hc[$l]
      if ($s.ProdInspTypeSettingIsActive -and ($l -notin $quyTacNhom) -and ($l -notin $GiuLai)) {
        $viec.Add([pscustomobject]@{ Ma=$m.Product; Nhom=$nhom; Loai=$l; HanhDong='TAT'; Etag=$s.'@odata.etag' })
      }
    }
  }
}
if ($GioiHan -gt 0) { $viec = @($viec | Select-Object -First $GioiHan) } else { $viec = $viec.ToArray() }
Write-Host ("Việc cần làm: {0} (TẠO: {1}, SỬA DM1: {2}, TẮT: {3}, KÍCH HOẠT: {4}) | Lô: {5}" -f $viec.Count,
  @($viec | ? HanhDong -eq 'TAO').Count, @($viec | ? HanhDong -eq 'SUA_DM1').Count,
  @($viec | ? HanhDong -eq 'TAT').Count, @($viec | ? HanhDong -eq 'KICH_HOAT').Count, $KichThuocLo) -ForegroundColor Cyan

if ($viec.Count -eq 0) { Write-Host 'Không có việc - thoát, không tạo log.' -ForegroundColor Green; return }
if (-not (Get-Module -ListAvailable ImportExcel)) {
  Write-Host 'Cài module ImportExcel...'
  Install-Module ImportExcel -Scope CurrentUser -Force
}

# ---- 2. Thực hiện ----
$log = New-Object System.Collections.Generic.List[object]
$sw = [Diagnostics.Stopwatch]::StartNew()
if ($ThucHien) { Get-Csrf }
for ($i = 0; $i -lt $viec.Count; $i += $KichThuocLo) {
  $lo = @($viec[$i..([math]::Min($i + $KichThuocLo, $viec.Count) - 1)])
  $kqLo = if ($ThucHien) { Send-Batch $lo } else { @($lo | % { @{ OK=$null; Code=''; Msg='DRY-RUN' } }) }
  for ($j = 0; $j -lt $lo.Count; $j++) {
    $v = $lo[$j]; $kq = $kqLo[$j]
    $log.Add([pscustomobject]@{ 'Thời gian'=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); 'Mã'=$v.Ma; 'Nhóm'=$v.Nhom
      'Loại kiểm'=$v.Loai; 'Hành động'=$v.HanhDong
      'Kết quả'= $(if ($kq.OK -eq $null) { 'DRY-RUN' } elseif ($kq.OK) { 'OK' } else { 'LỖI' })
      'HTTP'="$($kq.Code)"; 'Thông báo'=$kq.Msg })
  }
  $loi = @($kqLo | ? { $_.OK -eq $false }).Count
  if ($ThucHien) {
    Write-Host ("Lô {0}-{1}/{2}: OK {3}, LỖI {4} | {5:n0}s" -f ($i+1), ($i+$lo.Count), $viec.Count,
      ($lo.Count - $loi), $loi, $sw.Elapsed.TotalSeconds) -ForegroundColor $(if ($loi) { 'Yellow' } else { 'Green' })
  } else {
    Write-Host ("Lô {0}-{1}/{2}: DRY-RUN (chưa ghi)" -f ($i+1), ($i+$lo.Count), $viec.Count) -ForegroundColor Gray
  }
}

# ---- 3. Log ----
$tag = if ($ThucHien) { 'ThucHien' } else { 'DryRun' }
$out = Join-Path $PSScriptRoot ("Log_TaoInspSetup_{0}_{1}.xlsx" -f $tag, (Get-Date -Format 'yyyy-MM-dd_HHmm'))
$log | Export-Excel -Path $out -WorksheetName 'Log' -TableStyle Medium9 -AutoSize -FreezeTopRow -NoNumberConversion '*'
$soLoi = @($log | ? 'Kết quả' -eq 'LỖI').Count
Write-Host ("Xong: {0} | OK: {1} | LỖI: {2} | Tổng thời gian: {3:n0}s" -f $out,
  @($log | ? 'Kết quả' -eq 'OK').Count, $soLoi, $sw.Elapsed.TotalSeconds)
if ($soLoi -gt 0) { exit 1 }   # báo đỏ trên GitHub Actions để nhận email khi có dòng lỗi
