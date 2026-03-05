Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:results = @()
$script:jsonCount = 0

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

                # Barcha identifikator va kodlarni Excel avtomatik raqam deb o'ylamasligi uchun
                # PowerShell darajasida qat'iy [string] qilib olamiz.
                $fakturaId = [string]$data.facturaid
                $fakturaNo = [string]$data.facturadoc.facturano
                $fakturaDate = [string]$data.facturadoc.facturadate
                $contractNo = [string]$data.contractdoc.contractno
                $contractDate = [string]$data.contractdoc.contractdate
                
                $sellerName = [string]$data.seller.name
                $sellerStir = [string]$data.sellertin
                $sellerVat = [string]$data.seller.vatregcode
                $sellerAcc = [string]$data.seller.account
                $sellerMfo = [string]$data.seller.bankid
                $sellerAddr = [string]$data.seller.address
                $sellerDir = [string]$data.seller.director
                $sellerAccnt = [string]$data.seller.accountant
                
                $buyerName = [string]$data.buyer.name
                $buyerStir = [string]$data.buyertin
                $buyerVat = [string]$data.buyer.vatregcode
                $buyerAcc = [string]$data.buyer.account
                $buyerMfo = [string]$data.buyer.bankid
                $buyerAddr = [string]$data.buyer.address
                $buyerDir = [string]$data.buyer.director
                $buyerAccnt = [string]$data.buyer.accountant

                $products = $data.productlist.products

                if ($null -ne $products -and $products.Count -gt 0) {
                    foreach ($product in $products) {
                        
                        $katalogCode = [string]$product.catalogcode
                        
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
# Excelda "@" formati qat'iy matn degani. Bu orqali 20 xonali H/R lar nollarga aylanib ketmaydi
# va STIR/QQS/Katalog kodlar E+ ko'rinishiga kelmaydi.
# Ustunlar: Hujjat(2), Shartnoma(4), Sotuvchi STIR(7), Sotuvchi QQS(8), Sotuvchi H/R(9), 
#           Xaridor STIR(15), Xaridor QQS(16), Xaridor H/R(17), Katalog kodi(24)
$textCols = 2, 4, 7, 8, 9, 15, 16, 17, 24
foreach ($col in $textCols) {
    Set-ExcelColumn -Worksheet $sheet -Column $col -NumberFormat "@"
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