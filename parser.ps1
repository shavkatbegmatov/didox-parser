Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:results = @()
$script:jsonCount = 0

$zipFiles = Get-ChildItem -Path . -Filter "*.zip"

if ($zipFiles.Count -eq 0) {
    Write-Host "Ushbu papkada ZIP fayllar topilmadi." -ForegroundColor Yellow
    exit
}

# 1. YANADA MUSTAHKAMLANGAN FORMATLASH FUNKSIYALARI
$formatPul = {
    param($val)
    if ($null -eq $val -or $val -eq "") { return "0,00" }
    try {
        # Matn bo'lsa ham majburan sof matematik raqamga (Double) o'tkazamiz
        $num = [Convert]::ToDouble($val, [System.Globalization.CultureInfo]::InvariantCulture)
        return $num.ToString("0.00", [System.Globalization.CultureInfo]::InvariantCulture) -replace '\.', ','
    } catch {
        # Favqulodda holatda (xato bersa) faqat matndagi nuqtani almashtiramiz
        return [string]$val -replace '\.', ','
    }
}

$formatSoni = {
    param($val)
    if ($null -eq $val -or $val -eq "") { return "0" }
    try {
        $num = [Convert]::ToDouble($val, [System.Globalization.CultureInfo]::InvariantCulture)
        return $num.ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture) -replace '\.', ','
    } catch {
        return [string]$val -replace '\.', ','
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

                $fakturaId = $data.facturaid
                $fakturaNo = $data.facturadoc.facturano
                $fakturaDate = $data.facturadoc.facturadate
                $contractNo = $data.contractdoc.contractno
                $contractDate = $data.contractdoc.contractdate
                
                $sellerName = $data.seller.name
                $sellerStir = $data.sellertin
                $sellerVat = $data.seller.vatregcode
                $sellerAcc = $data.seller.account
                $sellerMfo = $data.seller.bankid
                $sellerAddr = $data.seller.address
                $sellerDir = $data.seller.director
                $sellerAccnt = $data.seller.accountant
                
                $buyerName = $data.buyer.name
                $buyerStir = $data.buyertin
                $buyerVat = $data.buyer.vatregcode
                $buyerAcc = $data.buyer.account
                $buyerMfo = $data.buyer.bankid
                $buyerAddr = $data.buyer.address
                $buyerDir = $data.buyer.director
                $buyerAccnt = $data.buyer.accountant

                $products = $data.productlist.products

                if ($null -ne $products -and $products.Count -gt 0) {
                    foreach ($product in $products) {
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
                            "Katalog Kodi" = $product.catalogcode
                            "Katalog Nomi" = $product.catalogname
                            "O'lchov Birligi" = $product.packagename
                            
                            "Soni" = &$formatSoni $product.count
                            "Narxi" = &$formatPul $product.summa
                            "Yetkazib Berish Narxi (QQSsiz)" = &$formatPul $product.deliverysum
                            "QQS Stavkasi (%)" = $product.vatrate
                            "QQS Summasi" = &$formatPul $product.vatsum
                            "Jami Summa (QQS bilan)" = &$formatPul $product.deliverysumwithvat
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

# Maxsus ustunlarni Excel buzib yubormasligi uchun sof MATN ("@") tipiga o'tkazish
# Hujjat(2), Shartnoma(4), STIR(7, 15), QQS(8, 16), H/R(9, 17), MFO(10, 18), Katalog kodi(24)
$textCols = 2, 4, 7, 8, 9, 10, 15, 16, 17, 18, 24
foreach ($col in $textCols) {
    Set-ExcelColumn -Worksheet $sheet -Column $col -NumberFormat "@"
}

Close-ExcelPackage $excel

Write-Host "Zo'r! Jami $script:jsonCount ta JSON fayl '$exportPath' ga muvaffaqiyatli saqlandi." -ForegroundColor Green