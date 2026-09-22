# Tao_InspPlan.ps1 - Tạo inspection plan cho các mã đã có loại kiểm nhưng chưa có plan.
# Mặc định DRY-RUN (chỉ lập kế hoạch, xuất Excel). -ThucHien để ghi.
# Nguồn giới hạn: field YY1 dạng "x +/- y" (đã đối chiếu với plan cũ). Không có -> để trống cho QA.
# Plan tạo ra ở STATUS 1 (chưa release).
param(
  [switch]$ThucHien,
  [int]$GioiHan = 0,
  [string]$ChiNhom = '',      # lọc theo nhóm mã, vd 'M052'
  [string]$ChiMa = '',        # chỉ làm 1 mã
  [switch]$GiuCm,             # NVL phụ: giữ nguyên đơn vị cm (mặc định đổi cm -> mm)
  [switch]$BoSung,            # vào cả những mã ĐÃ có plan để thêm chỉ tiêu còn thiếu
  [switch]$DongBo             # đồng bộ dung sai: trường YY1 đổi giá trị so với ảnh chụp lần trước -> sửa plan status 1
)
$ErrorActionPreference='Stop'; $ProgressPreference='SilentlyContinue'
$BaseUrl='https://my412079-api.s4hana.cloud.sap'
$SvcPlan="$BaseUrl/sap/opu/odata/sap/API_INSPECTIONPLAN_SRV"
$SvcPrd="$BaseUrl/sap/opu/odata/sap/API_PRODUCT_SRV"
$V4="$BaseUrl/sap/opu/odata4/sap/api_product/srvd_a2x/sap/product/0003"
$Plant='1000'
$CredFile=Join-Path $PSScriptRoot 'sap_cred.xml'

$NhomMa=[ordered]@{ 'M020'=@(20000000,20999999); 'M021'=@(21000000,21999999); 'M041'=@(41000000,41999999)
  'M042'=@(42000000,42999999); 'M044'=@(44000000,44999999); 'M050'=@(50000000,50999999)
  'M051'=@(51000000,51999999); 'M052'=@(52000000,52999999) }

# ---------- ĐỊNH NGHĨA CHỈ TIÊU ----------
# QT = định lượng; Field = field YY1 làm nguồn; DS = dung sai mặc định khi YY1 không ghi
$MIC=@{
  'NQ'   =@{T='Ngoại quan';QT=$false}
  'MS'   =@{T='Màu sắc';QT=$false}
  'VS'   =@{T='Vệ sinh';QT=$false}
  'DG'   =@{T='Đóng gói';QT=$false;Code='ĐG'}
  'CN'   =@{T='Chức năng';QT=$false}
  'CL'   =@{T='Chất liệu';QT=$false}
  'MU'   =@{T='Mùi';QT=$false}
  'DSTL' =@{T='Dung sai trọng lượng';QT=$true; Field='YY1_TrongLuong1_PRD'}
  'DSCD' =@{T='Dung sai chiều dài';QT=$true;  Field='YY1_ChieuDaimm1_PRD'; FieldNVL='YY1_DGTcmCDai_PRD';  Cm=$true; DS=1}
  'DSCR' =@{T='Dung sai chiều rộng';QT=$true; FieldNVL='YY1_DGTcmCRong_PRD'; Cm=$true; DS=1}
  'DSCC' =@{T='Dung sai chiều cao';QT=$true;  FieldNVL='YY1_DGTcmCCao_PRD';  Cm=$true; DS=1}
  'DSDKN'=@{T='Dung sai đường kính ngoài';QT=$true; Field='YY1_DKNmm1_PRD'; DS=1; Code='DSĐKN'}
  'DSDKT'=@{T='Dung sai đường kính trong';QT=$true; Field='YY1_DKTmm1_PRD'; DS=1; Code='DSĐKT'}
  'DSKM' =@{T='Dung sai khổ màng';QT=$true;   Field='YY1_KhoCuonMang1_PRD'; DS=1}
  'DSDDMC'=@{T='Dung sai độ dày màng cán';QT=$true; Field='YY1_KhoCuonMang_PRD'; Code='DSĐDMC'}
  'DSDD' =@{T='Dung sai độ dày';QT=$true; Field='YY1_DoDayPcsmm1_PRD'; FieldNVL='YY1_DoDayTuiperMang1_PRD'; QuyDoi='DoDay'; Code='DSĐD'}
  'DSDL' =@{T='Dung sai định lượng';QT=$true; Code='DSĐL'}
  'DSDKL'=@{T='Dung sai đường kính lõi';QT=$true; Code='DSĐKL'}
  'DSKC' =@{T='Dung sai khoảng cách giữa 1 bước ống';QT=$true}
  'DSKCMDO'=@{T='Dung sai khoảng cách màng đến ống';QT=$true}
  'DSCDT'=@{T='Dung sai chiều dài phần thân';QT=$true}
  'DSCDH'=@{T='Dung sai chiều dài phần hút';QT=$true}
  'DSCDOG'=@{T='Dung sai chiều dài ống gấp';QT=$true}
  'DSCCNM'=@{T='Dung sai chiều cao ngửa muỗng';QT=$true}
  'DSCRTC'=@{T='Dung sai chiều rộng tay cầm';QT=$true}
  'DSCRPM'=@{T='Dung sai chiều rộng phần muỗng';QT=$true}
}
function MICCode($k){ if ($MIC[$k].Code) { $MIC[$k].Code } else { $k } }

$MacDinh   = @('NQ','MS','VS','DG','CN','CL')
$BoNVL     = @('NQ','MS','VS','DG','CL','MU')
$BoNVLPhuB = @('NQ','VS','DG','CL','MU')
$BoXuatBan = @('NQ','VS','MS','DG','CN','DSTL')      # usage 6

# Product Group -> chi tieu them (usage 5)
$TheoGroup=@{
  '100001'=@('DSDKN','DSDKT','DSCD','DSTL'); '100002'=@('DSDKN','DSDKT','DSCD','DSTL')
  '100003'=@('DSDKN','DSDKT','DSCD','DSTL'); '100004'=@('DSDKN','DSDKT','DSCD','DSTL')
  '100006'=@('DSDKN','DSDKT','DSCD','DSTL'); '100007'=@('DSDKN','DSDKT','DSCD','DSTL')
  '100008'=@('DSDKN','DSDKT','DSCD','DSTL')
  '100005'=@('DSDKN','DSDKT','DSCD','DSTL','DSCRPM')      # ống muỗng
  '100009'=@('DSKM','DSKC')                                # TKI
  '100010'=@('DSKM','DSKC','DSKCMDO','DSCDT','DSCDH','DSCDOG')  # TKU
  '200001'=@('DSCD','DSCR','DSCC','DSTL','DSKM','DSDDMC'); '200002'=@('DSCD','DSCR','DSCC','DSTL','DSKM','DSDDMC')
  '200003'=@('DSDKN','DSDKT','DSCC','DSTL'); '200004'=@('DSDKN','DSDKT','DSCC','DSTL')
  '200005'=@('DSDKN','DSDKT','DSCC','DSTL'); '200006'=@('DSDKN','DSDKT','DSCC','DSTL')
  '200008'=@('DSDKN','DSDKT','DSCC','DSTL'); '200010'=@('DSDKN','DSDKT','DSCC','DSTL')
  '200011'=@('DSDKN','DSDKT','DSCC','DSTL'); '200012'=@('DSDKN','DSDKT','DSCC','DSTL')
  '200013'=@('DSDKN','DSDKT','DSCC','DSTL')
  '200009'=@('DSCD','DSCR','DSCCNM','DSCRTC','DSTL','DSDDMC','DSKM')   # muỗng định hình
  '200007'=@('DSKM','DSDDMC','DSDKL','DSTL')                            # màng HIPS
  '300002'=@('DSKM','DSDDMC')                                           # BTP màng
}
$GroupNVLChinh = @('400001','400002','400003','400004','400005','400006','400007','400008')
$GroupNVLPhuB  = @('500002','500004','500005','500006','500007','500008','500011','500012','500013','500014','500023','800047')
# Khong kiem trong luong: cuon mang HD/PE, muong nhua, decal, tem, giay decal, mang POF, cuon PP det, dua
$GroupNVLPhuA_KhongTL = @('500021','500022','500039','500009','600001','600002','500043','500042','200014')
$GroupBoQua    = @('300001')   # BTP keo: không lập plan
# Màng/giấy khổ lớn: có thể mang Product Group 4xxxx hoặc 5xxxx nhưng vẫn là NVL chính (M020)
# Tui (ke ca bao det 500001): co kiem DO DAY, khong kiem chieu cao
$GroupTui  = @('500001','500016','500017','500026','500027','500028','500029','500030','500031','500032','500033','500034','500035','500043')
# Tem / giay decal: KHONG kiem do day, KHONG kiem chieu cao
$GroupTem  = @('600001','600002')
$GroupMangGiay = @('400001','400003','400004','400006','500003','500020','500021','500022','500040','500042','500043')

if ($env:SAP_USER -and $env:SAP_PASS) { $pair="{0}:{1}" -f $env:SAP_USER,$env:SAP_PASS }
else { $c=Import-Clixml $CredFile; $pair="{0}:{1}" -f $c.UserName,$c.GetNetworkCredential().Password }
$auth='Basic '+[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
$hdrJ=@{Authorization=$auth;Accept='application/json'}
$sw=[Diagnostics.Stopwatch]::StartNew()
function Buoc($t){ Write-Host ("[{0,5:n0}s] {1}" -f $sw.Elapsed.TotalSeconds,$t) -ForegroundColor Cyan }
function GetV2($svc,$set,$select,$filter){ $all=New-Object System.Collections.Generic.List[object]; $skip=0; $top=5000
  do { $q="`$top=$top&`$skip=$skip&`$format=json"; if($select){$q="`$select=$select&$q"}; if($filter){$q="`$filter=$filter&$q"}
    $r=Invoke-RestMethod -Uri "$svc/$set`?$q" -Headers $hdrJ
    foreach($z in $r.d.results){$all.Add($z)}; $skip+=$top } while ($r.d.results.Count -eq $top)
  return $all }
function GetV4($path){ $all=New-Object System.Collections.Generic.List[object]; $skip=0; $top=5000
  do { $sep=if($path -match '\?'){'&'}else{'?'}
    $r=Invoke-RestMethod -Uri "$V4/$path$sep`$top=$top&`$skip=$skip" -Headers $hdrJ
    foreach($z in $r.value){$all.Add($z)}; $skip+=$top } while ($r.value.Count -eq $top)
  return $all }
function N($p){ "$p".TrimStart('0') }   # chuan hoa ma: bo so 0 dem dau
function NhomCua($p){ $n=0L; if(-not [long]::TryParse("$p".TrimStart('0'),[ref]$n)){return $null}
  foreach($k in $NhomMa.Keys){ if($n -ge $NhomMa[$k][0] -and $n -le $NhomMa[$k][1]){return $k} }; return $null }
function ParseYY1($s){ $t="$s".Trim(); if($t -eq ''){return $null}; $t=$t -replace ',','.'
  $m=[regex]::Match($t,'^\s*(-?\d+(?:\.\d+)?)\s*(?:\+/-|\+-|±)\s*(\d+(?:\.\d+)?)')
  if($m.Success){ return @{N=[double]$m.Groups[1].Value; T=[double]$m.Groups[2].Value} }
  $m2=[regex]::Match($t,'^\s*(-?\d+(?:\.\d+)?)\s*$')
  if($m2.Success){ $v=[double]$m2.Groups[1].Value; if($v -eq 0){return $null}; return @{N=$v; T=$null} }
  return $null }

# ---------- ĐỌC DỮ LIỆU ----------
Buoc 'Đọc master data'
$prd=GetV2 $SvcPrd 'A_Product' 'Product,ProductGroup,BaseUnit' $null
$desc=GetV2 $SvcPrd 'A_ProductDescription' 'Product,ProductDescription' $null
$pp=GetV4 "ProductPlant?`$filter=Plant eq '$Plant'&`$select=Product,IsMarkedForDeletion"
$setup=GetV4 "ProductPlantInspTypeSetting?`$filter=Plant eq '$Plant'&`$select=Product,InspectionLotType,ProdInspTypeSettingIsActive"
Buoc 'Đọc plan hiện có'
$hdrRaw=GetV2 $SvcPlan 'A_InspectionPlan' 'InspectionPlanGroup,InspectionPlan,IsDeleted,IsMarkedForDeletion,BillOfOperationsUsage,BillOfOperationsStatus,InspectionPlanInternalVersion,BillOfOperationsUnit' "Plant eq '$Plant'"
$assRaw=GetV2 $SvcPlan 'A_InspPlanMaterialAssgmt' 'Material,InspectionPlanGroup,InspectionPlan,IsDeleted' "Plant eq '$Plant'"

$gr=@{}; $bu=@{}; foreach($p in $prd){ $gr[(N $p.Product)]=$p.ProductGroup; $bu[(N $p.Product)]=$p.BaseUnit }


$dsc=@{}; foreach($d in $desc){ if(-not $dsc[(N $d.Product)]){$dsc[(N $d.Product)]=$d.ProductDescription} }
# Danh sach loai tru: ma_bo_qua.txt (moi dong 1 ma, # la ghi chu) -> ma loi master data, ma test...
$loaiTru=@{}
$fLT=Join-Path $PSScriptRoot 'ma_bo_qua.txt'
if (Test-Path $fLT) { Get-Content $fLT -Encoding UTF8 | % { ($_ -split '#')[0].Trim() } | ? { $_ } | % { $loaiTru[(N $_)]=$true }
                      if ($loaiTru.Count) { Buoc ("Loai tru {0} ma theo ma_bo_qua.txt" -f $loaiTru.Count) } }
$song=@{}; foreach($z in $pp){ if(-not $z.IsMarkedForDeletion){$song[(N $z.Product)]=$true} }
$coLoai=@{}; foreach($s in $setup){ if($s.ProdInspTypeSettingIsActive){$coLoai[(N $s.Product)]=$true} }
$planOK=@{}; foreach($h in $hdrRaw){ if(-not($h.IsDeleted -or $h.IsMarkedForDeletion)){ $planOK["$($h.InspectionPlanGroup)|$($h.InspectionPlan)"]=$h } }
$coPlan=@{}; $planCua=@{}
foreach($a in $assRaw){ if($a.IsDeleted){continue}
  if($planOK.ContainsKey("$($a.InspectionPlanGroup)|$($a.InspectionPlan)")){
    $u=$planOK["$($a.InspectionPlanGroup)|$($a.InspectionPlan)"].BillOfOperationsUsage
    $coPlan["$(N $a.Material)|$u"]=$true
    $planCua["$(N $a.Material)|$u"]=$planOK["$($a.InspectionPlanGroup)|$($a.InspectionPlan)"] } }

# YY1 (thong so ky thuat) - doc tu Product API V4 ban 0002
# KHONG doc het 16.000 ma: chi doc YY1 cua nhung ma thuc su can tao plan (xem ben duoi).
# -> nhanh, luon lay so moi nhat, khong can file cache.
$V4b="$BaseUrl/sap/opu/odata4/sap/api_product/srvd_a2x/sap/product/0002"
$yy=@{}
[xml]$mt=(Invoke-WebRequest -Uri "$V4b/`$metadata" -Headers @{Authorization=$auth;Accept='application/xml'} -UseBasicParsing).Content
$nsx=New-Object Xml.XmlNamespaceManager $mt.NameTable; $nsx.AddNamespace('edm','http://docs.oasis-open.org/odata/ns/edm')
$fl=@(); $ent=''
foreach($t in $mt.SelectNodes('//edm:EntityType',$nsx)){
  $f=@($t.SelectNodes('edm:Property',$nsx) | ? { $_.Name -match '^YY1_' })
  if ($f.Count -gt $fl.Count) { $fl=@($f | % Name); $ent=$t.Name } }
$setYY=$ent -replace '_Type$',''
$maGoc=@{}; foreach($p in $prd){ $maGoc[(N $p.Product)]="$($p.Product)" }

function DocYY1($dsMa){
  $ds=@($dsMa | Sort-Object -Unique)
  for($i=0; $i -lt $ds.Count; $i+=30){
    $lo=@($ds[$i..([math]::Min($i+29,$ds.Count-1))])
    $loc=($lo | % { "Product eq '$($maGoc[$_])'" }) -join ' or '
    $u="$V4b/$setYY`?`$select=Product,$($fl -join ',')&`$filter=$([uri]::EscapeDataString($loc))"
    $r=Invoke-RestMethod -Uri $u -Headers $hdrJ
    foreach($z in $r.value){ $yy[(N $z.Product)]=$z } } }

# ---------- LẬP KẾ HOẠCH ----------
Buoc 'Lập kế hoạch'
function BoChiTieu($ma){
  $g="$($gr[$ma])"; $nhom=NhomCua $ma
  if ($g -in $GroupBoQua) { return $null }
  if ($nhom -eq 'M042') {   # BTP gia cong: theo ten ma
    $ten="$($dsc[$ma])"
    if ($ten -match '(?i)\d+\s*mic') { return @($MacDinh + @('DSKM','DSDD','DSDKL','DSTL')) }
    if ($ten -match '(?i)\d+\s*gsm') { return @($MacDinh + @('DSKM','DSDL','DSDKL','DSTL')) }
    return @($BoNVL + @('DSCD','DSCR','DSDD','DSKM'))     # như cuộn màng HD
  }
  # Ma BO cua thanh pham (ten co dau '+'): chi kiem them trong luong
  if ($nhom -in @('M050','M051','M052') -and "$($dsc[$ma])" -match '\+') { return @($MacDinh + @('DSTL')) }
  if ($nhom -eq 'M020') {          # NVL chinh: xet NHOM MA truoc Product Group
    # mang/giay kho lon (dù Product Group nam trong dai 5xxxx) van la NVL chinh -> them kho mang + do day
    if ($g -in $GroupMangGiay) { return @($BoNVL + @('DSKM','DSDD')) }
    return $BoNVL }
  if ($g -in $GroupNVLChinh) { return $BoNVL }
  if ($g -in $GroupNVLPhuB)  { return $BoNVLPhuB }
  if ($g -eq '500024') {          # ong giay: dai, DK ngoai, DK trong, do day, trong luong
    return @($BoNVL + @('DSCD','DSDKN','DSDKT','DSDD','DSTL')) }
  if ($g -like '5*' -or $g -in $GroupTem -or $g -eq '200014') {
    $b=@($BoNVL + @('DSCD','DSCR'))
    if ($g -notin $GroupTem -and $g -notin $GroupTui) { $b += 'DSCC' }   # tem, tui: khong kiem chieu cao
    if ($g -in $GroupTui)    { $b += 'DSDD' }        # chi tui moi kiem do day
    if ($g -notin $GroupNVLPhuA_KhongTL) { $b += 'DSTL' }
    return $b }
  if ($g -eq '300003') {          # BTP ma le: nhu hang dinh hinh, suy dang tu TEN ma
    $ten="$($dsc[$ma])"
    if ($ten -match '(?i)tròn|tô|nắp|ly|hũ|chậu|thau|dĩa|nĩa') { return @($MacDinh + @('DSDKN','DSDKT','DSCC','DSTL')) }
    return @($MacDinh + @('DSCD','DSCR','DSCC','DSTL','DSKM','DSDDMC')) }
  if ($TheoGroup.ContainsKey($g)) { return @($MacDinh + $TheoGroup[$g]) }
  return $null     # chưa có quy tắc -> bỏ qua, ghi vào log
}
# Method cua chi tieu DINH LUONG (theo thong ke tren plan cu). Chi tieu dinh tinh: de MIC master tu cap.
$MethodDL = @{ 'DSTL'='TL'; 'DSCD'='KT'; 'DSCR'='KT'; 'DSCC'='KT'; 'DSDKN'='VS'; 'DSDKT'='CC'
  'DSKM'='KT'; 'DSDDMC'='ĐĐD'; 'DSDD'='ĐĐD'; 'DSDL'='TL'; 'DSDKL'='KT'; 'DSKC'='KT'; 'DSKCMDO'='KT'
  'DSCDT'='KT'; 'DSCDH'='KT'; 'DSCDOG'='KT'; 'DSCCNM'='KT'; 'DSCRTC'='KT'; 'DSCRPM'='KT' }
$DungSaiPhanTram = 0.05     # khong co dung sai ghi san va khong co mac dinh -> +/-5%
# Do day tui/bao tren ma: "3 ZEM" = 0.03mm | "25MIC" = 0.025mm | so thuan "0.3" = 0.03mm (chia 10)
function ParseDoDay($raw){
  $t="$raw".Trim(); if($t -eq ''){ return $null }
  $t=$t -replace ',','.'
  $m=[regex]::Match($t,'(-?\d+(?:\.\d+)?)\s*(?:(?:\+/-|\+-|±)\s*(\d+(?:\.\d+)?))?\s*([A-Za-zµ]*)')
  if(-not $m.Success){ return $null }
  $n=[double]$m.Groups[1].Value; if($n -eq 0){ return $null }
  $dv=$m.Groups[3].Value.ToUpper()
  $he = if ($dv -match '^ZEM') { 0.01 } elseif ($dv -match '^(MIC|MICRON|UM)') { 0.001 } else { 0.1 }   # so thuan: chia 10
  $t2=$null; if($m.Groups[2].Success){ $t2=[double]$m.Groups[2].Value * $he }
  return @{ N=[math]::Round($n*$he,3); T=$(if($t2 -ne $null){[math]::Round($t2,3)}); DonVi=$(if($dv){$dv}else{'(so thuan)'}) } }

function LaNVLPhu($ma){ $g="$($gr[$ma])"; return ($g -like '5*' -or $g -in $GroupTem -or $g -eq '200014') }
function LayGioiHan($ma,$chiTieu){
  # LUU Y: PowerShell khong phan biet hoa/thuong -> KHONG dat ten tham so la $mic (se che mat bang $MIC)
  $dn=$MIC[$chiTieu]
  if (-not $dn -or -not $dn.QT -or -not $yy.ContainsKey($ma)) { return $null }
  $f=$dn.Field; $doiCm=$false; $doiDoDay=$false
  if ((LaNVLPhu $ma) -and $dn.FieldNVL) {
    $f=$dn.FieldNVL; $doiCm=($dn.Cm -and -not $GiuCm); $doiDoDay=($dn.QuyDoi -eq 'DoDay') }
  if (-not $f) { return $null }
  $raw=$yy[$ma].($f)
  $p = if ($doiDoDay) { ParseDoDay $raw } else { ParseYY1 $raw }
  if (-not $p) { return $null }
  $n=$p.N; $t=$p.T
  if ($doiCm) { $n=$n*10; if ($t) { $t=$t*10 } }      # cm -> mm
  $ng="$f"
  if ($doiCm)    { $ng="$ng (cm->mm)" }
  if ($doiDoDay) { $ng="$ng ($($p.DonVi) ->mm)" }
  if ($t) { $ng="$ng +/- ghi san" }
  else {
    # dung sai rieng cho thung 3/5 lop
    if ("$($gr[$ma])" -in @('500036','500037') -and $chiTieu -in @('DSCD','DSCR','DSCC')) {
      $t = $(if ($chiTieu -eq 'DSCC') { 3 } else { 2 }); $ng="$ng + dung sai thung" }
    elseif ($dn.DS) { $t=$dn.DS; $ng="$ng + dung sai mac dinh" }
    else { $t=[math]::Round([math]::Abs($n)*$DungSaiPhanTram,3); $ng="$ng + dung sai 5%" } }
  # MIC chi cho 3 so thap phan -> lam tron tat ca
  return @{ Tgt=[math]::Round($n,3); L=[math]::Round($n-$t,3); U=[math]::Round($n+$t,3); Nguon=$ng } }

$viec=New-Object System.Collections.Generic.List[object]
$boQua=New-Object System.Collections.Generic.List[object]
# Loc truoc danh sach ma can xu ly (cung dieu kien voi vong lap ben duoi) -> chi doc YY1 cho nhom nay
$canDoc=New-Object System.Collections.Generic.List[string]
foreach($m0 in $coLoai.Keys){
  $m=N $m0
  if (-not $song[$m] -or $loaiTru[$m]) { continue }
  $nh=NhomCua $m; if (-not $nh) { continue }
  if ($ChiNhom -and $nh -ne $ChiNhom) { continue }
  if ($ChiMa -and $m -ne $ChiMa.TrimStart('0')) { continue }
  $thieu = $BoSung -or -not $coPlan["$m|5"] -or ($nh -in @('M050','M051','M052') -and -not $coPlan["$m|6"])
  if ($thieu) { $canDoc.Add($m) } }
Buoc ("Doc YY1 cho {0} ma can xu ly" -f $canDoc.Count)
if ($canDoc.Count) { DocYY1 $canDoc }
Buoc ("Da doc YY1: {0} ma" -f $yy.Count)

foreach($ma in ($coLoai.Keys | Sort-Object)) {
  $ma=N $ma
  if (-not $song[$ma] -or $loaiTru[$ma]) { continue }
  $nhom=NhomCua $ma; if (-not $nhom) { continue }
  if ($ChiNhom -and $nhom -ne $ChiNhom) { continue }
  if ($ChiMa -and $ma.TrimStart('0') -ne $ChiMa.TrimStart('0')) { continue }
  $bo=BoChiTieu $ma
  if (-not $bo) { $boQua.Add([pscustomobject]@{ 'Ma'=$ma;'Nhom'=$nhom;'Product Group'=$gr[$ma];'Tên'=$dsc[$ma];'Lý do'='Chưa có quy tắc / không lập plan' }); continue }
  $ds=@( @{U='5'; Bo=$bo} )
  if ($nhom -in @('M050','M051','M052')) { $ds += @{U='6'; Bo=$BoXuatBan} }
  foreach($d in $ds) {
    if ($coPlan["$ma|$($d.U)"] -and -not $BoSung) { continue }
    $ct=New-Object System.Collections.Generic.List[object]; $no=10
    # Khu trung lap: cac bo chi tieu cong don ($MacDinh + @(...)) co the lap lai 1 MIC
    # -> SAP bao QP/452 "Inspection characteristic already exists" va huy ca chi tieu do.
    $daDung=@{}
    foreach($k in $d.Bo) {
      if ($daDung[$k]) { continue }
      $daDung[$k]=$true
      $g=if($d.U -eq '5'){ LayGioiHan $ma $k } else { if($k -eq 'DSTL'){ LayGioiHan $ma $k } else { $null } }
      $ct.Add([pscustomobject]@{ So=("{0}" -f $no); MIC=(MICCode $k); Key=$k; Text=$MIC[$k].T; QT=$MIC[$k].QT
        Tgt=$(if($g){$g.Tgt}); L=$(if($g){$g.L}); U=$(if($g){$g.U}); Nguon=$(if($g){$g.Nguon}) })
      $no+=10 }
    $viec.Add([pscustomobject]@{ Ma=$ma; Nhom=$nhom; Grp=$gr[$ma]; Ten=$dsc[$ma]; Unit=$bu[$ma]; Usage=$d.U; CT=$ct })
  } }
if ($GioiHan -gt 0) { $viec=@($viec | Select-Object -First $GioiHan) } else { $viec=$viec.ToArray() }
Buoc ("Kế hoạch: {0} plan | {1} mã bỏ qua" -f $viec.Count,$boQua.Count)
$khop=@($viec | % { $_.Ma } | Sort-Object -Unique | ? { $yy.ContainsKey($_) }).Count
Buoc ("Ma co du lieu YY1: {0}/{1}" -f $khop, @($viec | % { $_.Ma } | Sort-Object -Unique).Count)

# ---------- XUẤT KẾ HOẠCH ----------
$stamp=Get-Date -Format 'yyyy-MM-dd_HHmm'
$keHoach = foreach($v in $viec){ foreach($c in $v.CT){
  [pscustomobject]@{ 'Ma'=$v.Ma;'Nhom'=$v.Nhom;'Product Group'=$v.Grp;'Tên mã'=$v.Ten;'Đơn vị'=$v.Unit;'Usage'=$v.Usage
    'So'=$c.So;'Chỉ tiêu'=$c.MIC;'Mô tả'=$c.Text;'Định lượng'=$(if($c.QT){'x'})
    'Target'=$c.Tgt;'LSL'=$c.L;'USL'=$c.U;'Nguồn giá trị'=$c.Nguon } } }
# Ten file kem nhom + ma tien trinh -> chay nhieu tab khong de len nhau
$hau = (@($ChiNhom,$ChiMa,("pid"+$PID)) | ? { $_ } ) -join '_'
$file=Join-Path $PSScriptRoot ("KeHoach_InspPlan_${stamp}_$hau.xlsx")
$x=@{Path=$file;TableStyle='Medium9';AutoSize=$true;FreezeTopRow=$true;NoNumberConversion='*'}
$tomTat = $viec | Group-Object Nhom,Usage | % { [pscustomobject]@{ 'Nhom'=$_.Group[0].Nhom;'Usage'=$_.Group[0].Usage;'Số plan'=$_.Count
  'Số chỉ tiêu TB'=[math]::Round(($_.Group | % { $_.CT.Count } | Measure-Object -Average).Average,1)
  'Có giới hạn'=@($_.Group | % { $_.CT } | ? { $_.Tgt -ne $null }).Count } }
$tomTat  | Export-Excel @x -WorksheetName 'Tom_tat'
$keHoach | Export-Excel @x -WorksheetName 'Ke_hoach'
if ($boQua.Count) { $boQua | Export-Excel @x -WorksheetName 'Bo_qua' }
Buoc "Kế hoạch: $file"
if (-not $ThucHien -and -not $DongBo) { Write-Host 'DRY-RUN - chưa ghi gì vào SAP. Thêm -ThucHien để tạo thật.' -ForegroundColor Green; return }
if (-not $ThucHien) { $viec=@() }   # -DongBo xem truoc: bo qua phan tao plan

# ---------- THỰC HIỆN ----------
$r=Invoke-WebRequest -Uri "$SvcPlan/`$metadata" -Headers @{Authorization=$auth;'x-csrf-token'='Fetch';Accept='application/xml'} -SessionVariable ss -UseBasicParsing
$csrf=$r.Headers['x-csrf-token']
$vsd="/Date(1735689600000)/"; $ved="/Date(253402214400000)/"
$script:msCuoi=0; $script:msDoc=0; $script:soDoc=0
function Post($set,$body){
  $sw=[Diagnostics.Stopwatch]::StartNew()
  try { $res=Invoke-WebRequest -Uri "$SvcPlan/$set" -Method Post -WebSession $ss `
      -Body ([Text.Encoding]::UTF8.GetBytes(($body|ConvertTo-Json -Compress))) `
      -Headers @{Authorization=$auth;'x-csrf-token'=$csrf;Accept='application/json'} -ContentType 'application/json' -UseBasicParsing
    $script:msCuoi=$sw.ElapsedMilliseconds
    return @{OK=$true; D=($res.Content|ConvertFrom-Json).d; Msg=''} }
  catch { $m=''
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $m=$_.ErrorDetails.Message }
    if (-not $m) { try { $st=$_.Exception.Response.GetResponseStream(); $m=(New-Object IO.StreamReader($st,[Text.Encoding]::UTF8)).ReadToEnd() } catch {} }
    $txt=''
    try { $j=$m|ConvertFrom-Json
      $er=@($j.error.innererror.errordetails | ? { $_.severity -eq 'error' } | % { "[$($_.code)] $($_.message)" })
      if ($er.Count) { $txt=$er -join ' ; ' } elseif ($j.error.message.value) { $txt="[$($j.error.code)] $($j.error.message.value)" } } catch { $txt=$m }
    $script:msCuoi=$sw.ElapsedMilliseconds
    return @{OK=$false; D=$null; Msg=$txt} } }

function XoaIfMatch(){
  # PowerShell nho lai header vao WebSession -> If-Match con lai se lam hong cac POST sau
  # (loi /IWFND/CM_MGW/537: eTag handling not supported for http method 'POST')
  try { if ($ss -and $ss.Headers -and $ss.Headers.ContainsKey('If-Match')) { [void]$ss.Headers.Remove('If-Match') } } catch {} }

function Patch($url,$body){
  $sw=[Diagnostics.Stopwatch]::StartNew()
  try { Invoke-WebRequest -Uri $url -Method Patch -WebSession $ss `
      -Body ([Text.Encoding]::UTF8.GetBytes(($body|ConvertTo-Json -Compress))) `
      -Headers @{Authorization=$auth;'x-csrf-token'=$csrf;Accept='application/json';'If-Match'='*'} `
      -ContentType 'application/json' -UseBasicParsing | Out-Null
    XoaIfMatch
    $script:msCuoi=$sw.ElapsedMilliseconds
    return @{OK=$true; Msg=''} }
  catch {
    XoaIfMatch
    $m=''
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $m=$_.ErrorDetails.Message }
    if (-not $m) { try { $st=$_.Exception.Response.GetResponseStream(); $m=(New-Object IO.StreamReader($st,[Text.Encoding]::UTF8)).ReadToEnd() } catch {} }
    $txt=''
    try { $j=$m|ConvertFrom-Json
      $er=@($j.error.innererror.errordetails | ? { $_.severity -eq 'error' } | % { "[$($_.code)] $($_.message)" })
      if ($er.Count) { $txt=$er -join ' ; ' } elseif ($j.error.message.value) { $txt="[$($j.error.code)] $($j.error.message.value)" } } catch { $txt=$m }
    $script:msCuoi=$sw.ElapsedMilliseconds
    return @{OK=$false; Msg=$txt} } }

# ---------- COPY MODEL CUA MIC (InspSpecTransferType) ----------
# '1' = Complete Copy Model : SAP chep du control indicator -> CHI duoc PATCH gia tri
# ''  = Incomplete Copy Model: SAP khong chep gi        -> phai tu gui du control indicator
# 'X' = Reference Characteristic: plan tro thang ve master -> KHONG PATCH gi ngoai optional
$SvcMIC='https://my412079-api.s4hana.cloud.sap/sap/opu/odata/sap/API_MASTERINSPCHARACTERISTIC_SRV'
$CopyModel=@{}
try {
  $mm=(Invoke-RestMethod -Uri "$SvcMIC/A_InspectionSpecification?`$filter=InspectionSpecificationPlant eq '$Plant'&`$top=5000&`$format=json" -Headers $hdrJ).d.results
  foreach($m in $mm){ $CopyModel[$m.InspectionSpecification]="$($m.InspSpecTransferType)" }
  $n1=@($CopyModel.Values | ? { $_ -eq '1' }).Count
  $nX=@($CopyModel.Values | ? { $_ -eq 'X' }).Count
  $n0=@($CopyModel.Values | ? { $_ -eq '' }).Count
  Buoc ("Copy model MIC: {0} Complete / {1} Reference / {2} Incomplete" -f $n1,$nX,$n0)
} catch { Write-Host "Khong doc duoc copy model MIC, se tu do theo phan hoi cua SAP" -ForegroundColor DarkYellow }

# Bang tra plan theo group, dung lai du lieu da doc 1 lan o dau script
$planTheoGroup=@{}
foreach($h in $hdrRaw){
  $k="$($h.InspectionPlanGroup)"
  if(-not $planTheoGroup[$k]){ $planTheoGroup[$k]=New-Object System.Collections.Generic.List[object] }
  $planTheoGroup[$k].Add($h) }

function DocSAP($url){
  $sw=[Diagnostics.Stopwatch]::StartNew()
  $r=(Invoke-RestMethod -Uri $url -Headers $hdrJ).d.results
  $script:msDoc += $sw.ElapsedMilliseconds; $script:soDoc++
  $script:msCuoi=$sw.ElapsedMilliseconds
  return $r }


# ---------- GOI GOP ($batch) ----------
# Moi request nam trong 1 changeset RIENG -> loi cua cai nay khong huy cai kia.
# Tra ve mang @{OK;Msg;D} theo dung thu tu da gui.
function BatchGui($reqs){
  if (-not $reqs -or $reqs.Count -eq 0) { return @() }
  $sw=[Diagnostics.Stopwatch]::StartNew()
  $bnd='batch_ttk'; $sb=New-Object Text.StringBuilder
  $i=0
  foreach($r in $reqs){
    $i++; $cs="changeset_$i"
    [void]$sb.AppendLine("--$bnd")
    [void]$sb.AppendLine("Content-Type: multipart/mixed; boundary=$cs")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("--$cs")
    [void]$sb.AppendLine('Content-Type: application/http')
    [void]$sb.AppendLine('Content-Transfer-Encoding: binary')
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("$($r.Method) $($r.Url) HTTP/1.1")
    [void]$sb.AppendLine('Content-Type: application/json')
    [void]$sb.AppendLine('Accept: application/json')
    if ($r.Method -ne 'POST') { [void]$sb.AppendLine('If-Match: *') }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine(($r.Body | ConvertTo-Json -Compress))
    [void]$sb.AppendLine("--$cs--") }
  [void]$sb.AppendLine("--$bnd--")
  $kq=@()
  try {
    $res=Invoke-WebRequest -Uri "$SvcPlan/`$batch" -Method Post -WebSession $ss `
      -Body ([Text.Encoding]::UTF8.GetBytes($sb.ToString())) `
      -Headers @{Authorization=$auth;'x-csrf-token'=$csrf;Accept='multipart/mixed'} `
      -ContentType "multipart/mixed; boundary=$bnd" -UseBasicParsing
    XoaIfMatch
    # tach tung phan hoi theo thu tu (chap nhan ca 'HTTP/1.1 201' lan 'HTTP/1.1  201')
    $phan=@([regex]::Split($res.Content,'(?m)^HTTP/1\.\d\s+') | Select-Object -Skip 1)
    if ($phan.Count -eq 0) {
      # Khong nhan dang duoc -> ghi nguyen phan hoi ra file de xem tan mat
      $dump=Join-Path $PSScriptRoot ("Batch_PhanHoi_{0}.txt" -f (Get-Date -Format 'HHmmss'))
      try { [IO.File]::WriteAllText($dump,$res.Content,[Text.Encoding]::UTF8)
            Write-Host "   !! Khong doc duoc batch, da ghi phan hoi ra: $dump" -ForegroundColor Red } catch {} }
    foreach($ph in $phan){
      $ma=0; if ($ph -match '^(\d{3})') { $ma=[int]$matches[1] }
      $body=''
      $k=$ph.IndexOf('{'); if ($k -ge 0) { $body=$ph.Substring($k) }
      $j=$null; try { $j=($body -replace '(?s)--.*$','') | ConvertFrom-Json } catch {}
      if ($ma -ge 200 -and $ma -lt 300) { $kq += @{OK=$true; Msg=''; D=$(if($j){$j.d}) } }
      else {
        $txt="HTTP $ma"
        try { $er=@($j.error.innererror.errordetails | ? { $_.severity -eq 'error' } | % { "[$($_.code)] $($_.message)" })
              if ($er.Count) { $txt=$er -join ' ; ' } elseif ($j.error.message.value) { $txt="[$($j.error.code)] $($j.error.message.value)" } } catch {}
        $kq += @{OK=$false; Msg=$txt; D=$null} } }
  } catch {
    $m=''; if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $m=$_.ErrorDetails.Message }
    foreach($r in $reqs){ $kq += @{OK=$false; Msg="batch loi: $m"; D=$null} } }
  $script:msCuoi=$sw.ElapsedMilliseconds
  while ($kq.Count -lt $reqs.Count) { $kq += @{OK=$false; Msg='khong doc duoc phan hoi batch'; D=$null} }
  return $kq }

$log=New-Object System.Collections.Generic.List[object]
$i=0
foreach($v in $viec){
  $i++; $grp=$v.Ma; $planNo='1'; $planMoi=$false
  Write-Host ("[{0}/{1}] {2} usage {3} ({4} chỉ tiêu)" -f $i,$viec.Count,$v.Ma,$v.Usage,$v.CT.Count)
  # 1. header
  $exist=@(if($planTheoGroup[$grp]){$planTheoGroup[$grp]})
  # Moi plan co the co nhieu phien ban noi bo (xoa trong app = them phien ban moi mang co xoa)
  # -> chi xet PHIEN BAN MOI NHAT cua tung plan
  $moiNhat=@($exist | Group-Object InspectionPlan | % {
    $_.Group | Sort-Object @{Expression={[int]("0"+$_.InspectionPlanInternalVersion)}} -Descending | Select-Object -First 1 })
  $sameUsage=@($moiNhat | ? { $_.BillOfOperationsUsage -eq $v.Usage -and -not $_.IsDeleted -and -not $_.IsMarkedForDeletion })
  if ($sameUsage.Count) {
    $planNo=$sameUsage[0].InspectionPlan; $planVer=$sameUsage[0].InspectionPlanInternalVersion
    # Don vi plan phai trung don vi co so vat tu, neu khong gan vat tu se loi CP/284 mai mai
    if ("$($sameUsage[0].BillOfOperationsUnit)" -and "$($sameUsage[0].BillOfOperationsUnit)" -ne "$($v.Unit)") {
      $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=0;Buoc='Header';KQ='LOI'
        Msg="plan $planNo dung don vi $($sameUsage[0].BillOfOperationsUnit) nhung don vi co so vat tu la $($v.Unit) -> xoa plan $planNo de script tao lai"}); continue }
    # Plan da release -> khong sua, de QA tu quyet dinh (tranh doi noi dung dang dung)
    if ("$($sameUsage[0].BillOfOperationsStatus)" -eq '4') {
      $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc='Bo qua';KQ='OK'
        Msg="plan $planNo da release (status 4) - khong sua"}); continue } }
  else {
    # Plan da xoa van chiem so thu tu trong SAP (loi CPCL/006) -> thu so tiep theo, toi da 9 lan
    $h=$null
    $max=0; foreach($e in @($exist)){ $n=0; if([int]::TryParse($e.InspectionPlan,[ref]$n)){ if($n -gt $max){$max=$n} } }
    $batDau = $max + 1
    for ($no=$batDau; $no -le ($batDau+20); $no++) {
      $planNo="$no"
      $h=Post 'A_InspectionPlan' ([ordered]@{ InspectionPlanGroup=$grp; InspectionPlan=$planNo; Plant=$Plant
        BillOfOperationsDesc="Plan $($v.Ma) U$($v.Usage)"; BillOfOperationsUsage=$v.Usage; BillOfOperationsStatus='1'
        MinimumLotSizeQuantity='0'; MaximumLotSizeQuantity='99999999'; BillOfOperationsUnit=$v.Unit
        ValidityStartDate=$vsd; ValidityEndDate=$ved })
      if ($h.OK) { break }
      if ("$($h.Msg)" -notmatch 'CPCL/006') { break }      # loi khac -> dung lai
      Write-Host "   plan $planNo da ton tai, thu so tiep theo" -ForegroundColor DarkGray }
    if (-not $h.OK) { $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc='Header';KQ='LOI';Msg=$h.Msg}); continue }
    $planVer=$h.D.InspectionPlanInternalVersion; $planMoi=$true
    $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc='Header';KQ='OK';Msg="plan=$planNo"}) }
  # 2. operation
  $op=$null
  if (-not $planMoi) {
    $ops=(DocSAP "$SvcPlan/A_InspPlanOperation?`$filter=InspectionPlanGroup eq '$grp' and InspectionPlan eq '$planNo' and Plant eq '$Plant'&`$format=json")
    # Nhieu ban ghi -> phai lay ban MOI NHAT, neu khong: CQCL/008 "cannot be uniquely assigned"
    $op=@($ops | ? { $_.Operation -eq '0010' -and -not $_.IsDeleted } |
          Sort-Object @{Expression={[int]("0"+$_.InspectionPlanInternalVersion)}},
                      @{Expression={[int]("0"+$_.BOOOpInternalVersionCounter)}} -Descending) | Select-Object -First 1 }
  if (-not $op) {
    $o=Post 'A_InspPlanOperation' ([ordered]@{ InspectionPlanGroup=$grp; InspectionPlan=$planNo; Plant=$Plant
      Operation='0010'; OperationText='Kiểm tra chất lượng'; OperationControlProfile='QM01'
      OperationReferenceQuantity='1'; OperationUnit=$v.Unit; OpQtyToBaseQtyNmrtr='1'; OpQtyToBaseQtyDnmntr='1'
      ValidityStartDate=$vsd; ValidityEndDate=$ved })
    if (-not $o.OK) { $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc='Operation';KQ='LOI';Msg=$o.Msg}); continue }
    # BAT BUOC doc lai operation: phan hoi POST khong mang du khoa phien ban,
    # dung thang se loi CQCL/008 "cannot be uniquely assigned to one operation".
    $ops=(DocSAP "$SvcPlan/A_InspPlanOperation?`$filter=InspectionPlanGroup eq '$grp' and InspectionPlan eq '$planNo' and Plant eq '$Plant'&`$format=json")
    $op=@($ops | ? { $_.Operation -eq '0010' -and -not $_.IsDeleted } |
          Sort-Object @{Expression={[int]("0"+$_.InspectionPlanInternalVersion)}},
                      @{Expression={[int]("0"+$_.BOOOpInternalVersionCounter)}} -Descending) | Select-Object -First 1
    $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc='Operation';KQ='OK';Msg="opId=$($op.BOOOperationInternalID)"}) }
  if (-not $op) { $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc='Operation';KQ='LOI';Msg='Không đọc được operation'}); continue }
  # 3. chỉ tiêu
  $daCo=@{}; $daCoSo=@()
  if (-not $planMoi) { (DocSAP "$SvcPlan/A_InspPlanOpCharacteristic?`$filter=InspectionPlanGroup eq '$grp' and InspectionPlan eq '$planNo'&`$format=json") |
    ? { -not $_.IsDeleted } | % { $daCo[$_.InspectionSpecification]=$true
                                  $z=0; if([int]::TryParse($_.BOOCharacteristic,[ref]$z)){ $daCoSo+=$z } } }
  # So thu tu bat dau: sau chi tieu lon nhat dang co (tranh QP/452 khi bo sung)
  $soMax=0
  if (-not $planMoi) { foreach($k in $daCoSo){ if($k -gt $soMax){$soMax=$k} } }
  # LUU Y: KHONG dung $batch cho phan nay. SAP xu ly moi changeset nhu 1 phien nap
  # task list rieng -> khoa operation khong con hop le -> CQCL/008 + CPCC_DM/002.
  $n=$soMax
  foreach($c in $v.CT){
    if ($daCo[$c.MIC]) { continue }
    $n+=10
    $b=[ordered]@{ InspectionPlanGroup=$grp; InspectionPlan=$planNo
      InspectionPlanInternalVersion=$op.InspectionPlanInternalVersion
      BOOOperationInternalID=$op.BOOOperationInternalID; BOOOpInternalVersionCounter=$op.BOOOpInternalVersionCounter
      BOOCharacteristic=("{0}" -f $n)
      InspectionSpecification=$c.MIC; InspectionSpecificationVersion='1'; InspectionSpecificationPlant=$Plant
      InspectionSpecificationText=$c.Text
      InspCharacteristicSampleUnit=$v.Unit; BOOCharcSampleQuantity='1'
      SamplingProcedure=$(if($c.QT){ if($v.Usage -eq '6'){'TTK03L'}else{'TTK_AQLL'} } else { if($v.Usage -eq '6'){'TTK03'}else{'TTK_AQL'} })
      ValidityStartDate=$vsd; ValidityEndDate=$ved }
    $rr=Post 'A_InspPlanOpCharacteristic' $b
    $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc="Chỉ tiêu $($c.MIC)";KQ=$(if($rr.OK){'OK'}else{'LOI'});Msg=$rr.Msg})
    if (-not $rr.OK) { break }
    if ($rr.D) {
      $d=$rr.D
      $key = "InspectionPlanGroup='$($d.InspectionPlanGroup)',InspectionPlan='$($d.InspectionPlan)'," +
             "InspectionPlanInternalVersion='$($d.InspectionPlanInternalVersion)'," +
             "BOOOperationInternalID='$($d.BOOOperationInternalID)'," +
             "BOOOpInternalVersionCounter='$($d.BOOOpInternalVersionCounter)'," +
             "BOOCharacteristic='$($d.BOOCharacteristic)',BOOCharacteristicVersion='$($d.BOOCharacteristicVersion)'"
      $tt=$CopyModel[$c.MIC]; $laRef=($tt -eq 'X')
      if ($tt -eq '') { $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=0;Buoc="Canh bao $($c.MIC)";KQ='LOI'
        Msg='MIC con o Incomplete Copy Model - can sua ve Complete Copy trong MIC master'}) }
      $bp=[ordered]@{}; $ghi=''
      if ($c.QT -and -not $laRef) {
        if ($MethodDL[$c.Key]) { $bp.InspectionMethod=$MethodDL[$c.Key]
                                 $bp.InspectionMethodPlant=$Plant; $bp.InspectionMethodVersion='1' }
        $bp.InspSpecDecimalPlaces=3
        if ($c.Tgt -ne $null) {
          $bp.InspSpecTargetValue="$($c.Tgt)"; $bp.InspSpecLowerLimit="$($c.L)"; $bp.InspSpecUpperLimit="$($c.U)"
          $ghi="tgt=$($c.Tgt) lsl=$($c.L) usl=$($c.U)" }
        else { $ghi='chua co so lieu -> de trong, so le=3' } }
      if ($bp.Keys.Count -gt 0) {
        $pr=Patch "$SvcPlan/A_InspPlanOpCharacteristic($key)" $bp
        $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc="Sua $($c.MIC)";KQ=$(if($pr.OK){'OK'}else{'LOI'})
          Msg=$(if($pr.OK){$ghi.Trim()}else{$pr.Msg})}) }
      if ($v.Usage -eq '6') {
        $po=Patch "$SvcPlan/A_InspPlanOpCharacteristic($key)" ([ordered]@{ InspSpecCharcCategory='' })
        $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc="Optional $($c.MIC)";KQ=$(if($po.OK){'OK'}else{'LOI'})
          Msg=$(if($po.OK){'optional'}else{$po.Msg})}) } } }
  # 4. gán vật tư
  $ass=@(); if (-not $planMoi) { $ass=(DocSAP "$SvcPlan/A_InspPlanMaterialAssgmt?`$filter=Material eq '$($v.Ma)' and Plant eq '$Plant' and InspectionPlanGroup eq '$grp' and InspectionPlan eq '$planNo'&`$format=json") }
  if (-not @($ass).Count) {
    $a=Post 'A_InspPlanMaterialAssgmt' ([ordered]@{ Material=$v.Ma; Plant=$Plant; InspectionPlanGroup=$grp; InspectionPlan=$planNo
      ValidityStartDate=$vsd; ValidityEndDate=$ved })
    $log.Add([pscustomobject]@{Ma=$v.Ma;Usage=$v.Usage;Ms=$script:msCuoi;Buoc='Gán vật tư';KQ=$(if($a.OK){'OK'}else{'LOI'});Msg=$a.Msg}) }
  }

# ---------- ĐỒNG BỘ DUNG SAI (-DongBo) ----------
# Phat hien thay doi bang cach SO GIA TRI tung truong YY1 voi anh chup lan truoc
# (KHONG dung LastChangeDateTime: moc thoi gian vat tu nhay vi du thu ly do).
#   - Chi truong YY1 nao doi gia tri -> chi tieu lay nguon tu truong do moi duoc xet.
#   - Plan status 1: tu PATCH gia tri moi.  Plan status 4: chi ghi log de QA quyet.
#   - Lan dau chua co anh chup: chi tao moc, khong sua gi.
function ChuanYY($v){ $t="$v".Trim(); $d=0.0
  if ($t -and [double]::TryParse($t,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$d) -and $d -eq 0) { return '' }
  return $t }
if ($DongBo) {
  $fSnap=Join-Path $PSScriptRoot 'yy1_anh_chup.json'
  # Lay dung ten truong theo metadata ($select cua OData phan biet hoa/thuong)
  $tenChuan=@{}; foreach($x in $fl){ $tenChuan[$x.ToUpper()]=$x }
  $truong=@($MIC.Values | % { $_.Field; $_.FieldNVL } | ? { $_ } | % { $tenChuan["$_".ToUpper()] } | ? { $_ } | Sort-Object -Unique)
  $thieu=@($MIC.Values | % { $_.Field; $_.FieldNVL } | ? { $_ -and -not $tenChuan["$_".ToUpper()] } | Sort-Object -Unique)
  if ($thieu.Count) { Buoc ("Canh bao: khong tim thay truong trong API: {0}" -f ($thieu -join ', ')) }
  Buoc ("Dong bo: doc {0} truong YY1 lien quan cua toan bo vat tu" -f $truong.Count)
  $hienTai=@{}; $skip=0; $top=5000
  do { $r=Invoke-RestMethod -Uri "$V4b/$setYY`?`$select=Product,$($truong -join ',')&`$top=$top&`$skip=$skip" -Headers $hdrJ
       foreach($z in $r.value){ $m=N $z.Product; $o=@{}; foreach($t in $truong){ $o[$t]=ChuanYY $z.$t }
                                $hienTai[$m]=$o; if(-not $yy.ContainsKey($m)){ $yy[$m]=$z } }
       $skip+=$top } while ($r.value.Count -eq $top)
  Buoc ("Dong bo: da doc {0} vat tu" -f $hienTai.Count)

  $truoc=$null
  if (Test-Path $fSnap) { try { $truoc=Get-Content $fSnap -Raw -Encoding UTF8 | ConvertFrom-Json } catch {} }
  $luuAnh=$false
  if (-not $truoc) {
    Buoc 'Dong bo: chua co anh chup YY1 -> tao moc ban dau, tu lan sau moi so sanh'
    $luuAnh=$true }
  else {
    # ma -> danh sach truong da doi
    $doi=@{}
    foreach($m in $hienTai.Keys){
      $c=$truoc.$m; if (-not $c) { continue }          # ma moi: da co luong tao plan lo
      foreach($t in $truong){ if ((ChuanYY $c.$t) -ne $hienTai[$m][$t]) {
        if(-not $doi[$m]){ $doi[$m]=New-Object System.Collections.Generic.List[string] }
        $doi[$m].Add($t) } } }
    $dsDoi=@($doi.Keys | ? { ($coPlan["$_|5"] -or $coPlan["$_|6"]) -and -not $loaiTru[$_] })
    if ($ChiNhom) { $dsDoi=@($dsDoi | ? { (NhomCua $_) -eq $ChiNhom }) }
    Buoc ("Dong bo: {0} ma doi YY1, trong do {1} ma co plan" -f $doi.Count, $dsDoi.Count)
    foreach($ma in $dsDoi){
      $bo=@(BoChiTieu $ma)
      foreach($us in @('5','6')){
        $h=$planCua["$ma|$us"]; if (-not $h) { continue }
        $st="$($h.BillOfOperationsStatus)"; $ds=$null
        foreach($k in $bo){
          $dn=$MIC[$k]; if (-not $dn.QT) { continue }
          if ($us -eq '6' -and $k -ne 'DSTL') { continue }
          # truong nguon cua chi tieu nay (cung logic voi LayGioiHan)
          $f=$dn.Field; if ((LaNVLPhu $ma) -and $dn.FieldNVL) { $f=$dn.FieldNVL }
          if ($f) { $f=$tenChuan["$f".ToUpper()] }
          if (-not $f -or -not ($doi[$ma] -contains $f)) { continue }   # truong nguon khong doi -> bo qua
          $g=LayGioiHan $ma $k; if (-not $g) { continue }
          if ($ds -eq $null) { $ds=@(DocSAP "$SvcPlan/A_InspPlanOpCharacteristic?`$filter=InspectionPlanGroup eq '$($h.InspectionPlanGroup)' and InspectionPlan eq '$($h.InspectionPlan)'&`$format=json" | ? { -not $_.IsDeleted }) }
          $c=$ds | ? { $_.InspectionSpecification -eq (MICCode $k) } | Select-Object -First 1
          if (-not $c) { continue }
          $cu="$([double]("0"+$c.InspSpecTargetValue))/$([double]("0"+$c.InspSpecLowerLimit))/$([double]("0"+$c.InspSpecUpperLimit))"
          $moi="$([double]$g.Tgt)/$([double]$g.L)/$([double]$g.U)"
          $vcu="$($truoc.$ma.$f)"; $vmoi=$hienTai[$ma][$f]
          $ghi="$f`: '$vcu' -> '$vmoi' | plan $cu -> $moi"
          if ($cu -eq $moi) { continue }
          if ($st -eq '4') {
            $log.Add([pscustomobject]@{Ma=$ma;Usage=$us;Ms=0;Buoc="Dong bo $(MICCode $k)";KQ='CANH BAO';Msg="plan da release, KHONG sua. $ghi"}); continue }
          $key = "InspectionPlanGroup='$($c.InspectionPlanGroup)',InspectionPlan='$($c.InspectionPlan)'," +
                 "InspectionPlanInternalVersion='$($c.InspectionPlanInternalVersion)'," +
                 "BOOOperationInternalID='$($c.BOOOperationInternalID)'," +
                 "BOOOpInternalVersionCounter='$($c.BOOOpInternalVersionCounter)'," +
                 "BOOCharacteristic='$($c.BOOCharacteristic)',BOOCharacteristicVersion='$($c.BOOCharacteristicVersion)'"
          if ($ThucHien) {
            $pr=Patch "$SvcPlan/A_InspPlanOpCharacteristic($key)" ([ordered]@{ InspSpecDecimalPlaces=3
                  InspSpecTargetValue="$($g.Tgt)"; InspSpecLowerLimit="$($g.L)"; InspSpecUpperLimit="$($g.U)" })
            $log.Add([pscustomobject]@{Ma=$ma;Usage=$us;Ms=$script:msCuoi;Buoc="Dong bo $(MICCode $k)";KQ=$(if($pr.OK){'OK'}else{'LOI'})
              Msg=$(if($pr.OK){$ghi}else{$pr.Msg})}) }
          else { $log.Add([pscustomobject]@{Ma=$ma;Usage=$us;Ms=0;Buoc="Dong bo $(MICCode $k)";KQ='XEM TRUOC';Msg=$ghi}) } } } }
    # Chi cap nhat anh chup khi chay that va khong co loi (loi thi lan sau xet lai)
    if ($ThucHien -and -not @($log | ? { $_.KQ -eq 'LOI' -and "$($_.Buoc)" -like 'Dong bo*' }).Count) { $luuAnh=$true } }
  if ($luuAnh) {
    $giu=@{}; foreach($v in $viec){ $giu["$($v.Ma)"]=$true }
    $gon=@{}
    foreach($m in $hienTai.Keys){
      if (-not ($coPlan["$m|5"] -or $coPlan["$m|6"] -or $giu[$m])) { continue }
      $o=@{}; foreach($t in $truong){ if ($hienTai[$m][$t]) { $o[$t]=$hienTai[$m][$t] } }
      $gon[$m]=$o }
    try { $gon | ConvertTo-Json -Depth 3 -Compress | Set-Content $fSnap -Encoding UTF8
          Buoc ("Dong bo: da luu anh chup YY1 ({0} ma) -> {1}" -f $gon.Count,$fSnap) } catch { Buoc "Khong luu duoc anh chup: $($_.Exception.Message)" } } }

$fileLog=Join-Path $PSScriptRoot ("Log_TaoInspPlan_${stamp}_$hau.xlsx")
$log | Export-Excel -Path $fileLog -WorksheetName 'Log' -TableStyle Medium9 -AutoSize -FreezeTopRow -NoNumberConversion '*'

$tong = $log | Group-Object { ($_.Buoc -split ' ')[0] } | % {
  [pscustomobject]@{ 'Bước'=$_.Name; 'Số lần'=$_.Count
    'Tổng giây'=[math]::Round((($_.Group | Measure-Object Ms -Sum).Sum)/1000,1)
    'TB ms'=[math]::Round((($_.Group | Measure-Object Ms -Average).Average),0) } } | Sort-Object 'Tổng giây' -Descending
Write-Host "`n=== THOI GIAN THEO BUOC ===" -ForegroundColor Cyan
$tong | Format-Table -AutoSize
Write-Host ("Doc SAP trong vong lap: {0} lan, {1} giay" -f $script:soDoc,[math]::Round($script:msDoc/1000,1)) -ForegroundColor Yellow
try { $tong | Export-Excel -Path $fileLog -WorksheetName 'ThoiGian' -TableStyle Medium9 -AutoSize } catch {}

Buoc ("XONG. Log: {0} | OK: {1} | LỖI: {2}" -f $fileLog, @($log|? KQ -eq 'OK').Count, @($log|? KQ -eq 'LOI').Count)
if (@($log | ? { $_.KQ -eq 'LOI' }).Count) { exit 1 }
