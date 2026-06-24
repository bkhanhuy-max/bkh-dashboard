# extract_ai_data.ps1
# Xuất toàn bộ dữ liệu chi tiết từ các file Excel sang ai_data.json cho Trợ lý ảo AI

Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Resolve-NewestFile($dir, $pattern, $fallback) {
    if (Test-Path $dir) {
        $files = Get-ChildItem -Path $dir -Filter $pattern | Sort-Object LastWriteTime -Descending
        if ($files.Count -gt 0) {
            return $files[0].FullName
        }
    }
    return $fallback
}

$uniceDir = "d:\BKH-AI\01. DU AN 2026\01. UNICE\01. REPORT"
$howellDir = "d:\BKH-AI\01. DU AN 2026\02. HOWELL\01. REPORT"

$uniceBudgetFile   = Resolve-NewestFile $uniceDir "*NGAN SACH*.xlsx" "$uniceDir\1. UNICE - NGAN SACH - UP 15.06.2026.xlsx"
$uniceMaterialFile = Resolve-NewestFile $uniceDir "*THEO DOI VAT TU*.xlsx" "$uniceDir\2. UNICE - THEO DOI VAT TU - UP 15.06.2026.xlsx"
$unicePlanFile     = Resolve-NewestFile $uniceDir "*KE HOACH*.xlsx" "$uniceDir\3. UNICE - KE HOACH VT.NTP.xlsx"
$uniceVoFile       = Resolve-NewestFile $uniceDir "*PHAT SINH*.xlsx" "$uniceDir\4. UNICE - THEO DOI PHAT SINH TDTK.xlsx"
$howellBudgetFile  = Resolve-NewestFile $howellDir "*NGAN SACH*.xlsx" "$howellDir\1. HOWELL - NGAN SACH - UP 15.06.2026.xlsx"
$howellPlanFile    = Resolve-NewestFile $howellDir "*KE HOACH*.xlsx" "$howellDir\3. HOWELL - KE HOACH VT.NTP.xlsx"

function Get-SafeDouble($val, $fallback = 0.0) {
    if ($val -eq $null -or $val -eq "") { return $fallback }
    $out = 0.0
    if ([double]::TryParse($val.ToString().Trim(), [ref]$out)) { return $out }
    return $fallback
}
function Get-SafeText($val) {
    if ($val -eq $null) { return "" }
    return $val.ToString().Trim()
}
function Get-SafeDate($val) {
    if ($val -eq $null -or $val -eq "") { return "" }
    $out = 0.0
    if ([double]::TryParse($val.ToString(), [ref]$out)) {
        try {
            return [DateTime]::FromOADate($out).ToString("dd/MM/yyyy")
        } catch {
            return $val.ToString()
        }
    }
    return $val.ToString()
}

$ai = [PSCustomObject]@{
    last_updated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    unice = [PSCustomObject]@{
        contracts         = @()
        submittals        = @()
        variation_orders  = @()
        materials_detail  = @()
        concrete_all      = @()
        budget_categories = @()
        suppliers_all     = @()
    }
    howell = [PSCustomObject]@{
        budget_categories = @()
        boq_detail        = @()
        contracts         = @()
        materials         = @()
    }
}

# ─────────────────────────────────────────────
# UNICE FILE 3: KẾ HOẠCH - HỢP ĐỒNG & MẪU TRÌNH
# ─────────────────────────────────────────────
Write-Output "1. Đọc hợp đồng & mẫu trình UNICE..."
if (Test-Path $unicePlanFile) {
    try {
        $wb = $excel.Workbooks.Open($unicePlanFile, [System.Type]::Missing, $true)

        # Sheet: 3, KH KHĐ — Danh sách hợp đồng
        $sh = $wb.Sheets.Item("3, KH KHĐ")
        $contracts = @()
        for ($r = 12; $r -le 300; $r++) {
            $stt    = Get-SafeText $sh.Cells.Item($r, 1).Value2
            $name   = Get-SafeText $sh.Cells.Item($r, 2).Value2
            if ($name -eq "" -and $stt -eq "") { continue }
            if ($name -eq "") { continue }
            $contractor = Get-SafeText $sh.Cells.Item($r, 3).Value2
            $value      = Get-SafeDouble $sh.Cells.Item($r, 7).Value2
            $signed     = Get-SafeText $sh.Cells.Item($r, 8).Value2
            $status     = Get-SafeText $sh.Cells.Item($r, 13).Value2
            $note       = Get-SafeText $sh.Cells.Item($r, 14).Value2
            $contracts += [PSCustomObject]@{
                STT        = $stt
                TenHopDong = $name
                NhaThau    = $contractor
                GiaTri     = $value
                NgayKy     = $signed
                TrangThai  = $status
                GhiChu     = $note
            }
        }
        $ai.unice.contracts = $contracts
        Write-Output "   Hop dong: $($contracts.Count) dong"

        # Sheet: 4, MAU TRINH CDT — Mẫu trình chủ đầu tư
        $sh2 = $wb.Sheets.Item("4, MAU TRINH CDT")
        $vals = $sh2.UsedRange.Value2
        $rows = $vals.GetLength(0)
        $submittals = @()
        for ($r = 12; $r -le $rows; $r++) {
            $stt    = Get-SafeText $vals[$r, 1]
            $name   = Get-SafeText $vals[$r, 2]
            if ($name -eq "") { continue }
            $category = Get-SafeText $vals[$r, 3]
            $dateSubmit = Get-SafeText $vals[$r, 4]
            $dateApprove = Get-SafeText $vals[$r, 5]
            $status = Get-SafeText $vals[$r, 6]
            $note   = Get-SafeText $vals[$r, 7]
            $submittals += [PSCustomObject]@{
                STT         = $stt
                TenMauTrinh = $name
                HangMuc     = $category
                NgayTrinh   = $dateSubmit
                NgayDuyet   = $dateApprove
                TrangThai   = $status
                GhiChu      = $note
            }
        }
        $ai.unice.submittals = $submittals
        Write-Output "   Mau trinh CDT: $($submittals.Count) dong"

        $wb.Close($false)
    } catch {
        Write-Warning "Loi doc KE HOACH: $_"
    }
}

# ─────────────────────────────────────────────
# UNICE FILE 4: THEO DÕI PHÁT SINH (VO)
# ─────────────────────────────────────────────
Write-Output "2. Đọc phát sinh VO UNICE..."
if (Test-Path $uniceVoFile) {
    try {
        $wb = $excel.Workbooks.Open($uniceVoFile, [System.Type]::Missing, $true)
        $sh = $wb.Sheets.Item("1, THEO DOI")
        $vos = @()
        for ($r = 15; $r -le 400; $r++) {
            $stt    = Get-SafeText $sh.Cells.Item($r, 1).Value2
            $desc   = Get-SafeText $sh.Cells.Item($r, 7).Value2
            if ($desc -eq "" -and $stt -eq "") { continue }
            if ($desc -eq "" -and $stt -notmatch "^\d+$") { continue }
            $category   = Get-SafeText $sh.Cells.Item($r, 2).Value2
            $type       = Get-SafeText $sh.Cells.Item($r, 3).Value2
            $valueBoq   = Get-SafeDouble $sh.Cells.Item($r, 8).Value2
            $valueApprv = Get-SafeDouble $sh.Cells.Item($r, 9).Value2
            $status     = Get-SafeText $sh.Cells.Item($r, 10).Value2
            $note       = Get-SafeText $sh.Cells.Item($r, 11).Value2
            $vos += [PSCustomObject]@{
                STT         = $stt
                HangMuc     = $category
                LoaiPS      = $type
                MoTa        = $desc
                GiaTriBoq   = $valueBoq
                GiaTriDuyet = $valueApprv
                TrangThai   = $status
                GhiChu      = $note
            }
        }
        $ai.unice.variation_orders = $vos
        Write-Output "   VO: $($vos.Count) dong"
        $wb.Close($false)
    } catch {
        Write-Warning "Loi doc VO: $_"
    }
}

# ─────────────────────────────────────────────
# UNICE FILE 2: THEO DÕI VẬT TƯ
# ─────────────────────────────────────────────
Write-Output "3. Đọc vật tư & bê tông UNICE..."
if (Test-Path $uniceMaterialFile) {
    try {
        $wb = $excel.Workbooks.Open($uniceMaterialFile, [System.Type]::Missing, $true)

        # Tất cả sự kiện đổ bê tông
        $shBT = $wb.Sheets.Item("II.2 BT NHAP THUC TE")
        $concreteAll = @()
        for ($r = 6; $r -le 600; $r++) {
            $stt = $shBT.Cells.Item($r, 1).Text
            if (-not ($stt -and $stt.Trim() -match "^\d+$")) { continue }
            $date   = Get-SafeText $shBT.Cells.Item($r, 2).Value2
            $name   = Get-SafeText $shBT.Cells.Item($r, 3).Value2
            $design = Get-SafeDouble $shBT.Cells.Item($r, 5).Value2
            $actual = Get-SafeDouble $shBT.Cells.Item($r, 6).Value2
            $mac    = Get-SafeText $shBT.Cells.Item($r, 9).Value2
            if ($actual -gt 0 -or $design -gt 0) {
                $concreteAll += [PSCustomObject]@{
                    STT        = $stt
                    Ngay       = $date
                    HangMuc    = $name
                    KhoiLuongTK = $design
                    KhoiLuongTT = $actual
                    MacBetong  = $mac
                }
            }
        }
        $ai.unice.concrete_all = $concreteAll
        Write-Output "   Be tong: $($concreteAll.Count) su kien"

        # Tất cả vật tư chi tiết
        $shVT = $wb.Sheets.Item("BANG TONG HOP")
        $vals = $shVT.UsedRange.Value2
        $rows = $vals.GetLength(0)
        $mats = @()
        for ($r = 9; $r -le $rows; $r++) {
            $name = Get-SafeText $vals[$r, 3]
            if ($name -eq "") { continue }
            $calc      = Get-SafeDouble $vals[$r, 4]
            $used      = Get-SafeDouble $vals[$r, 6]
            $remaining = Get-SafeDouble $vals[$r, 8]
            $unit      = Get-SafeText $vals[$r, 2]
            $mats += [PSCustomObject]@{
                TenVatTu   = $name
                DonVi      = $unit
                KhoiLuongKH = $calc
                DaSuDung   = $used
                ConLai     = $remaining
            }
        }
        $ai.unice.materials_detail = $mats
        Write-Output "   Vat tu: $($mats.Count) loai"

        $wb.Close($false)
    } catch {
        Write-Warning "Loi doc VAT TU: $_"
    }
}

# ─────────────────────────────────────────────
# UNICE FILE 1: NGÂN SÁCH — Chi tiết hạng mục
# ─────────────────────────────────────────────
Write-Output "4. Đọc ngân sách chi tiết UNICE..."
if (Test-Path $uniceBudgetFile) {
    try {
        $wb = $excel.Workbooks.Open($uniceBudgetFile, [System.Type]::Missing, $true)
        $shNS = $wb.Sheets.Item("1. NGAN SACH")
        $vals = $shNS.UsedRange.Value2
        $rows = $vals.GetLength(0)
        $budgetCats = @()
        for ($r = 12; $r -le $rows; $r++) {
            $id   = Get-SafeText $vals[$r, 2]
            $name = Get-SafeText $vals[$r, 4]
            if ($name -eq "") { continue }
            $boq      = Get-SafeDouble $vals[$r, 6]
            $approved = Get-SafeDouble $vals[$r, 9]
            $updated  = Get-SafeDouble $vals[$r, 13]
            $spent    = Get-SafeDouble $vals[$r, 17]
            $remaining = Get-SafeDouble $vals[$r, 21]
            $budgetCats += [PSCustomObject]@{
                MaHangMuc  = $id
                TenHangMuc = $name
                GiaTriBoq  = $boq
                AngSachDuyet = $approved
                AngSachCapNhat = $updated
                DaChiTieu  = $spent
                ConLai     = $remaining
            }
        }
        $ai.unice.budget_categories = $budgetCats
        Write-Output "   Ngan sach hang muc: $($budgetCats.Count) dong"

        # Danh sách nhà thầu phụ từ sheet "0, CHI PHI"
        $shCP = $wb.Sheets.Item("0, CHI PHI")
        $suppliers = @()
        for ($r = 5; $r -le 3700; $r++) {
            $supplier = Get-SafeText $shCP.Cells.Item($r, 5).Value2
            $category = Get-SafeText $shCP.Cells.Item($r, 3).Value2
            $amount   = Get-SafeDouble $shCP.Cells.Item($r, 10).Value2
            $date     = Get-SafeText $shCP.Cells.Item($r, 1).Value2
            $desc     = Get-SafeText $shCP.Cells.Item($r, 7).Value2
            if ($supplier -eq "" -or $amount -le 0) { continue }
            $suppliers += [PSCustomObject]@{
                NhaThau   = $supplier
                HangMuc   = $category
                GiaTri    = $amount
                Ngay      = $date
                MoTa      = $desc
            }
        }
        $ai.unice.suppliers_all = $suppliers
        Write-Output "   Nha thau chi phi: $($suppliers.Count) dong"

        $wb.Close($false)
    } catch {
        Write-Warning "Loi doc NGAN SACH UNICE: $_"
    }
}

# ─────────────────────────────────────────────
# HOWELL FILE 1: NGÂN SÁCH
# ─────────────────────────────────────────────
Write-Output "5. Đọc ngân sách HOWELL..."
if (Test-Path $howellBudgetFile) {
    try {
        $wb = $excel.Workbooks.Open($howellBudgetFile, [System.Type]::Missing, $true)

        # Sheet ngân sách HOWELL
        $shNS = $wb.Sheets.Item("01. NGAN SACH")
        $vals = $shNS.UsedRange.Value2
        $rows = $vals.GetLength(0)
        $hwCats = @()
        for ($r = 12; $r -le $rows; $r++) {
            $id   = Get-SafeText $vals[$r, 2]
            $name = Get-SafeText $vals[$r, 4]
            if ($name -eq "") { continue }
            $boq      = Get-SafeDouble $vals[$r, 6]
            $approved = Get-SafeDouble $vals[$r, 9]
            $updated  = Get-SafeDouble $vals[$r, 13]
            $spent    = Get-SafeDouble $vals[$r, 17]
            $remaining = Get-SafeDouble $vals[$r, 21]
            $hwCats += [PSCustomObject]@{
                MaHangMuc  = $id
                TenHangMuc = $name
                GiaTriBoq  = $boq
                AngSachDuyet = $approved
                AngSachCapNhat = $updated
                DaChiTieu  = $spent
                ConLai     = $remaining
            }
        }
        $ai.howell.budget_categories = $hwCats
        Write-Output "   HOWELL ngan sach: $($hwCats.Count) dong"

        # Sheet BOQ
        try {
            $shBOQ = $wb.Sheets.Item("00, BOQ - DAU VAO")
            $boqVals = $shBOQ.UsedRange.Value2
            $boqRows = $boqVals.GetLength(0)
            $boqDetail = @()
            for ($r = 10; $r -le $boqRows; $r++) {
                $maBoq   = Get-SafeText $boqVals[$r, 1]
                $hangMuc = Get-SafeText $boqVals[$r, 2]
                $name    = Get-SafeText $boqVals[$r, 8]
                if ($name -eq "") { continue }
                $unit     = Get-SafeText $boqVals[$r, 9]
                $qty      = Get-SafeDouble $boqVals[$r, 10]
                $material = Get-SafeDouble $boqVals[$r, 11]
                $labor    = Get-SafeDouble $boqVals[$r, 12]
                $total    = Get-SafeDouble $boqVals[$r, 13]
                if ($total -le 0 -and $qty -le 0) { continue }
                $boqDetail += [PSCustomObject]@{
                    MaBoq      = $maBoq
                    HangMuc    = $hangMuc
                    TenCongTac = $name
                    DonVi      = $unit
                    KhoiLuong  = $qty
                    VatLieu    = $material
                    NhanCong   = $labor
                    ThanhTien  = $total
                }
            }
            $ai.howell.boq_detail = $boqDetail
            Write-Output "   HOWELL BOQ: $($boqDetail.Count) hang muc"
        } catch {
            Write-Warning "Khong doc duoc sheet BOQ HOWELL: $_"
        }

        $wb.Close($false)
    } catch {
        Write-Warning "Loi doc NGAN SACH HOWELL: $_"
    }
}

# ─────────────────────────────────────────────
# HOWELL FILE 3: KẾ HOẠCH (VẬT TƯ & HỢP ĐỒNG)
# ─────────────────────────────────────────────
Write-Output "6. Đọc kế hoạch vật tư & hợp đồng HOWELL..."
if (Test-Path $howellPlanFile) {
    try {
        $wb = $excel.Workbooks.Open($howellPlanFile, [System.Type]::Missing, $true)

        # 1. Đọc sheet index 1: KH TDVT (Kế hoạch trình mẫu vật tư)
        $sh1 = $wb.Sheets.Item(1)
        $materials = @()
        $vals1 = $sh1.UsedRange.Value2
        $rows1 = $vals1.GetLength(0)
        for ($r = 12; $r -le $rows1; $r++) {
            $stt = Get-SafeText $vals1[$r, 1]
            $name = Get-SafeText $vals1[$r, 3]
            if ($name -eq "") { continue }
            $specs = Get-SafeText $vals1[$r, 4]
            $origin = Get-SafeText $vals1[$r, 5]
            $dateUse = Get-SafeDate $vals1[$r, 6]
            $dateSubmit = Get-SafeDate $vals1[$r, 7]
            $status = Get-SafeText $vals1[$r, 10]
            
            $materials += [PSCustomObject]@{
                STT = $stt
                TenVatTu = $name
                ChungLoai = $specs
                NguonGoc = $origin
                NgaySuDung = $dateUse
                NgayTrinh = $dateSubmit
                TrangThai = $status
            }
        }
        $ai.howell.materials = $materials
        Write-Output "   HOWELL vat tu: $($materials.Count) dong"

        # 2. Đọc sheet index 2: KH KKHĐ (Hợp đồng thầu phụ & tổ đội)
        $sh2 = $wb.Sheets.Item(2)
        $contracts = @()
        $vals2 = $sh2.UsedRange.Value2
        $rows2 = $vals2.GetLength(0)
        for ($r = 12; $r -le $rows2; $r++) {
            $stt = Get-SafeText $vals2[$r, 1]
            $name = Get-SafeText $vals2[$r, 2]
            if ($name -eq "") { continue }
            # Bỏ qua dòng tiêu đề La Mã
            if ($stt.ToString().Trim() -match "^[IVXLC]+$") { continue }

            $dateSign = Get-SafeDate $vals2[$r, 10]
            $dateStart = Get-SafeDate $vals2[$r, 11]
            $duration = Get-SafeText $vals2[$r, 12]
            $dateEnd = Get-SafeDate $vals2[$r, 13]
            $contractor = Get-SafeText $vals2[$r, 14]
            $status = Get-SafeText $vals2[$r, 15]

            $contracts += [PSCustomObject]@{
                STT = $stt
                TenHangMuc = $name
                NhaThau = $contractor
                NgayKy = $dateSign
                NgayThiCong = $dateStart
                TienDo = $duration
                NgayHoanThanh = $dateEnd
                TrangThai = $status
            }
        }
        $ai.howell.contracts = $contracts
        Write-Output "   HOWELL hop dong: $($contracts.Count) dong"

        $wb.Close($false)
    } catch {
        Write-Warning "Loi doc KE HOACH HOWELL: $_"
    }
}

# ─────────────────────────────────────────────
# Dọn dẹp & xuất file
# ─────────────────────────────────────────────
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

$jsonPath = "d:\BKH-AI\ai_data.json"
$ai | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding utf8
$size = (Get-Item $jsonPath).Length / 1KB
Write-Output "Xuat thanh cong $jsonPath ($([math]::Round($size, 1)) KB)"
Write-Output "  - UNICE hop dong: $($ai.unice.contracts.Count)"
Write-Output "  - UNICE mau trinh: $($ai.unice.submittals.Count)"
Write-Output "  - UNICE VO phat sinh: $($ai.unice.variation_orders.Count)"
Write-Output "  - UNICE vat tu: $($ai.unice.materials_detail.Count)"
Write-Output "  - UNICE be tong: $($ai.unice.concrete_all.Count)"
Write-Output "  - UNICE ngan sach hang muc: $($ai.unice.budget_categories.Count)"
Write-Output "  - UNICE nha thau chi tiet: $($ai.unice.suppliers_all.Count)"
Write-Output "  - HOWELL ngan sach: $($ai.howell.budget_categories.Count)"
Write-Output "  - HOWELL BOQ: $($ai.howell.boq_detail.Count)"
Write-Output "  - HOWELL hop dong: $($ai.howell.contracts.Count)"
Write-Output "  - HOWELL vat tu: $($ai.howell.materials.Count)"

