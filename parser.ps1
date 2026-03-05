Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:results = @()
$script:jsonCount = 0

# JSON dan katta raqamlarni string sifatida olish uchun regex funksiya
# ConvertFrom-Json katta raqamlarni double ga o'giradi va precision yo'qoladi
# Shuning uchun H/R, vatregcode kabi maydonlarni regex bilan olamiz
function Get-JsonStringValue {
    param(
        [string]$JsonString,
        [string]$FieldName
    )
    # "fieldname" : "value" yoki "fieldname" : 12345 (ham string ham number)
    if ($JsonString -match """$FieldName""\s*:\s*""([^""]*?)""") {
        return $Matches[1]
    }
    elseif ($JsonString -match """$FieldName""\s*:\s*(\d[\d.]*)") {
        return $Matches[1]
    }
    return ""
}

$zipFiles = Get-ChildItem -Path . -Filter "*.zip"

if ($zipFiles.Count -eq 0) {
    Write-Host "Ushbu papkada ZIP fayllar topilmadi." -ForegroundColor Yellow
    exit
}

# 1. SUMMALARNI VA SONLARNI TO'G'RI FORMATLASH
# Buni skriptning o'zida emas, Excelning NumberFormat orqali qilganimiz ancha xavfsiz. 
# Shuning uchun bu funksiyalarda faqat sof raqamga o'giramiz.
$convertToNumber = {
    param($val)
    if ($null -eq $val -or $val -eq "") { return 0 }
    try {
        # Agar nuqta yoki vergul bo'lsa, invariant madaniyat bilan double ga o'tkazish
        $valStr = [string]$val -replace ',', '.'
        return [Convert]::ToDouble($valStr, [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        return 0
    }
}

function Read-ZipArchive {
    param ([System.IO.Stream]$Stream)

    try {
        $archive = New-Object System.IO.Compression.ZipArchive($Stream, [System.IO.Compression.ZipArchiveMode]::Read)
        
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName.EndsWith(".json", [System.StringComparison]::OrdinalIgnoreCase)) {
                $script:jsonCount++
                
                $entryStream = $entry.Open()
                $reader = New-Object System.IO.StreamReader($entryStream)
                $jsonString = $reader.ReadToEnd()
                $reader.Close()
                $entryStream.Close()

                $data = $jsonString | ConvertFrom-Json

                # Katta raqamli maydonlarni REGEX orqali to'g'ridan-to'g'ri string sifatida olamiz
                # Bu double precision yo'qolishini oldini oladi (H/R 20 xonali, vatregcode 12 xonali)
                $sellerVat = Get-JsonStringValue -JsonString $jsonString -FieldName "vatregcode"
                $sellerAcc = Get-JsonStringValue -JsonString $jsonString -FieldName "account"
                
                # seller va buyer alohida bo'lgani uchun, avval seller blokini ajratamiz
                # Oddiyroq yo'l: seller va buyer blokdan keyin qayta oqiymiz
                
                # Seller blokini topamiz
                $sellerBlock = ""
                if ($jsonString -match '"seller"\s*:\s*\{([^}]+)\}') {
                    $sellerBlock = $Matches[1]
                }
                $buyerBlock = ""
                if ($jsonString -match '"buyer"\s*:\s*\{([^}]+)\}') {
                    $buyerBlock = $Matches[1]
                }
                
                # Seller maydonlari (regex orqali - precision saqlanadi)
                $sellerVat = ""
                if ($sellerBlock -match '"vatregcode"\s*:\s*"?(\d+)"?') { $sellerVat = $Matches[1] }
                $sellerAcc = ""
                if ($sellerBlock -match '"account"\s*:\s*"?(\d+)"?') { $sellerAcc = $Matches[1] }
                
                # Buyer maydonlari (regex orqali)
                $buyerVat = ""
                if ($buyerBlock -match '"vatregcode"\s*:\s*"?(\d+)"?') { $buyerVat = $Matches[1] }
                $buyerAcc = ""
                if ($buyerBlock -match '"account"\s*:\s*"?(\d+)"?') { $buyerAcc = $Matches[1] }
                
                # Qolgan identifikatorlar (ConvertFrom-Json orqali, chunki ular qisqa yoki string)
                $fakturaId = [string]$data.facturaid
                $fakturaNo = [string]$data.facturadoc.facturano
                $fakturaDate = [string]$data.facturadoc.facturadate
                $contractNo = [string]$data.contractdoc.contractno
                $contractDate = [string]$data.contractdoc.contractdate
                
                # sellertin va buyertin (9-10 xonali, double muammo yo'q)
                $sellerStir = [string]$data.sellertin
                $buyerStir = [string]$data.buyertin
                
                $sellerName = [string]$data.seller.name
                $sellerMfo = [string]$data.seller.bankid
                $sellerAddr = [string]$data.seller.address
                $sellerDir = [string]$data.seller.director
                $sellerAccnt = [string]$data.seller.accountant
                
                $buyerName = [string]$data.buyer.name
                $buyerMfo = [string]$data.buyer.bankid
                $buyerAddr = [string]$data.buyer.address
                $buyerDir = [string]$data.buyer.director
                $buyerAccnt = [string]$data.buyer.accountant

                $products = $data.productlist.products

                if ($null -ne $products -and $products.Count -gt 0) {
                    foreach ($product in $products) {
                        
                        # catalogcode ham katta raqam bo'lishi mumkin - regex orqali olamiz
                        $katalogCode = [string]$product.catalogcode
                        # Agar catalogcode float ga aylanib ketgan bo'lsa, original JSON dan regex bilan olamiz
                        # Har bir product uchun alohida regex qilish murakkab, shuning uchun
                        # ConvertFrom-Json dan kelgan qiymatni tekshiramiz
                        if ($katalogCode -match 'E\+' -or $katalogCode -match '\.') {
                            # Float formatda kelgan, original JSONdan izlaymiz
                            # ordno orqali identifikatsiya qilamiz
                            $ordNo = [string]$product.ordno
                            if ($jsonString -match """ordno""\s*:\s*""?$ordNo""?\s*,[\s\S]*?""catalogcode""\s*:\s*""?(\d+)""?") {
                                $katalogCode = $Matches[1]
                            }
                        }
                        
                        # Summalarni sof raqam sifatida olamiz, Excel formatlaydi
                        $soni = &$convertToNumber $product.count
                        $narxi = &$convertToNumber $product.summa
                        $yetkazibBerish = &$convertToNumber $product.deliverysum
                        $qqsStavkasi = &$convertToNumber $product.vatrate
                        $qqsSummasi = &$convertToNumber $product.vatsum
                        $jamiSumma = &$convertToNumber $product.deliverysumwithvat

                        $row = [PSCustomObject]@{
                            "Hujjat ID" = $fakturaId
                            "Hujjat Raqami" = $fakturaNo
                            "Hujjat Sanasi" = $fakturaDate
                            "Shartnoma Raqami" = $contractNo
                            "Shartnoma Sanasi" = $contractDate
                            
                            "Sotuvchi Tashkilot" = $sellerName
                            "Sotuvchi STIR" = $sellerStir
                            "Sotuvchi QQS Kodi" = $sellerVat
                            "Sotuvchi H/R" = $sellerAcc
                            "Sotuvchi MFO" = $sellerMfo
                            "Sotuvchi Manzili" = $sellerAddr
                            "Sotuvchi Rahbari" = $sellerDir
                            "Sotuvchi Bosh Hisobchisi" = $sellerAccnt
                            
                            "Xaridor Tashkilot" = $buyerName
                            "Xaridor STIR" = $buyerStir
                            "Xaridor QQS Kodi" = $buyerVat
                            "Xaridor H/R" = $buyerAcc
                            "Xaridor MFO" = $buyerMfo
                            "Xaridor Manzili" = $buyerAddr
                            "Xaridor Rahbari" = $buyerDir
                            "Xaridor Bosh Hisobchisi" = $buyerAccnt
                            
                            "Mahsulot T/R" = $product.ordno
                            "Mahsulot Nomi" = $product.name
                            "Katalog Kodi" = $katalogCode
                            "Katalog Nomi" = $product.catalogname
                            "O'lchov Birligi" = $product.packagename
                            
                            "Soni" = $soni
                            "Narxi" = $narxi
                            "Yetkazib Berish Narxi (QQSsiz)" = $yetkazibBerish
                            "QQS Stavkasi (%)" = $qqsStavkasi
                            "QQS Summasi" = $qqsSummasi
                            "Jami Summa (QQS bilan)" = $jamiSumma
                        }
                        $script:results += $row
                    }
                }
            }
            elseif ($entry.FullName.EndsWith(".zip", [System.StringComparison]::OrdinalIgnoreCase)) {
                $innerStream = $entry.Open()
                $memStream = New-Object System.IO.MemoryStream
                $innerStream.CopyTo($memStream)
                $innerStream.Close()
                
                $memStream.Position = 0
                Read-ZipArchive -Stream $memStream
                $memStream.Dispose()
            }
        }
        $archive.Dispose()
    }
    catch {
        Write-Host "Xatolik yuz berdi: $_" -ForegroundColor Red
    }
}

Write-Host "Kengaytirilgan ma'lumotlar tahlil qilinmoqda..." -ForegroundColor Cyan

foreach ($zip in $zipFiles) {
    $fileStream = [System.IO.File]::OpenRead($zip.FullName)
    Read-ZipArchive -Stream $fileStream
    $fileStream.Close()
}

if ($script:jsonCount -eq 0) {
    Write-Host "Hech qanday JSON fayl topilmadi." -ForegroundColor Yellow
    exit
}

$exportPath = "Maksimal_Fakturalar_hisoboti.xlsx"

$excel = $script:results | Export-Excel -Path $exportPath -AutoSize -BoldTopRow -FreezeTopRow -PassThru
$sheet = $excel.Workbook.Worksheets[1]

# -----------------------------------------------------------------------------
# 1. MATNLI USTUNLARNI FORMATLASH (E+ VA NOLLARNI YO'QOTISH UCHUN)
# Export-Excel string qiymatlarni ham raqamga o'girib yuborishi mumkin.
# Shuning uchun EPPlus orqali har bir yacheykani qat'iy string qilib qayta yozamiz.
# Ustunlar: Hujjat Raqami(2), Shartnoma Raqami(4), Sotuvchi STIR(7), Sotuvchi QQS(8), 
#           Sotuvchi H/R(9), Xaridor STIR(15), Xaridor QQS(16), Xaridor H/R(17), Katalog Kodi(24)
$textCols = 2, 4, 7, 8, 9, 15, 16, 17, 24
$totalRows = $sheet.Dimension.End.Row

foreach ($col in $textCols) {
    # Format ni "@" (Text) qilamiz
    for ($row = 1; $row -le $totalRows; $row++) {
        $cell = $sheet.Cells[$row, $col]
        $cell.Style.Numberformat.Format = "@"
    }
}

# Endi DATA qatorlarini results dan qayta yozamiz (string sifatida)
# Bu eng ishonchli usul - EPPlus yacheykaga .Value = [string] qo'yilsa, matn sifatida saqlaydi
$colMap = @{
    2  = "Hujjat Raqami"
    4  = "Shartnoma Raqami"
    7  = "Sotuvchi STIR"
    8  = "Sotuvchi QQS Kodi"
    9  = "Sotuvchi H/R"
    15 = "Xaridor STIR"
    16 = "Xaridor QQS Kodi"
    17 = "Xaridor H/R"
    24 = "Katalog Kodi"
}

for ($i = 0; $i -lt $script:results.Count; $i++) {
    $rowNum = $i + 2  # 1-qator sarlavha
    $item = $script:results[$i]
    foreach ($col in $textCols) {
        $propName = $colMap[$col]
        $val = $item.$propName
        if ($null -ne $val) {
            $sheet.Cells[$rowNum, $col].Value = [string]$val
        }
    }
}

# -----------------------------------------------------------------------------
# 2. SUMMALAR VA MIQDORLAR UCHUN FORMATLASH
# Siz so'ragan ko'rinish: ### ### ### ###,##. 
# ExcelNumberFormat: "#,##0.00" (bu lokal sozlamaga qarab bo'shliq va vergul qilib beradi)
# Soni(27), Narxi(28), Yetkazib Berish Narxi (QQSsiz)(29), QQS Summasi(31), Jami Summa (QQS bilan)(32)
$moneyCols = 28, 29, 31, 32
foreach ($col in $moneyCols) {
    Set-ExcelColumn -Worksheet $sheet -Column $col -NumberFormat "#,##0.00"
}

# Soni ustunini biroz farqli - "#,##0.##" qilamiz (agar qoldiqsiz bo'lsa butun ko'rsatadi)
Set-ExcelColumn -Worksheet $sheet -Column 27 -NumberFormat "#,##0.##"

Close-ExcelPackage $excel

Write-Host "Zo'r! Jami $script:jsonCount ta JSON fayl '$exportPath' ga muvaffaqiyatli saqlandi." -ForegroundColor Green