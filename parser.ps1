Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:results = @()
$script:jsonCount = 0

$zipFiles = Get-ChildItem -Path . -Filter "*.zip"

if ($zipFiles.Count -eq 0) {
    Write-Host "Ushbu papkada ZIP fayllar topilmadi." -ForegroundColor Yellow
    exit
}

function Read-ZipArchive {
    param ([System.IO.Stream]$Stream)

    try {
        $archive = New-Object System.IO.Compression.ZipArchive($Stream, [System.IO.Compression.ZipArchiveMode]::Read)
        
        foreach ($entry in $archive.Entries) {
            # JSON topilsa
            if ($entry.FullName.EndsWith(".json", [System.StringComparison]::OrdinalIgnoreCase)) {
                $script:jsonCount++
                
                $entryStream = $entry.Open()
                $reader = New-Object System.IO.StreamReader($entryStream)
                $jsonString = $reader.ReadToEnd()
                $reader.Close()
                $entryStream.Close()

                $data = $jsonString | ConvertFrom-Json

                # Ma'lumotlarni xavfsiz o'qish (yo'q bo'lsa $null bo'lib qolaveradi, xato bermaydi)
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

                # Agar mahsulotlar ro'yxati mavjud bo'lsa
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
                            "Soni" = $product.count
                            "Narxi" = $product.summa
                            "Yetkazib Berish Narxi (QQSsiz)" = $product.deliverysum
                            "QQS Stavkasi (%)" = $product.vatrate
                            "QQS Summasi" = $product.vatsum
                            "Jami Summa (QQS bilan)" = $product.deliverysumwithvat
                        }
                        $script:results += $row
                    }
                }
            }
            # Yana ZIP arxiv topilsa, ichiga kirish (rekursiya)
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
        Write-Host "Faylni o'qishda kichik xatolik: $_" -ForegroundColor Red
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

# Ma'lumotlarni yozish
$excel = $script:results | Export-Excel -Path $exportPath -AutoSize -BoldTopRow -FreezeTopRow -PassThru
$sheet = $excel.Workbook.Worksheets[1]

# 1. STIR va QQS kodlari uchun oddiy raqam formati ("0")
$numberCols = 7, 8, 15, 16
foreach ($col in $numberCols) {
    Set-ExcelColumn -Worksheet $sheet -Column $col -NumberFormat "0"
}

# 2. H/R, MFO va Katalog kodlar ma'lumoti buzilmasligi uchun MATN formati ("@")
$textCols = 9, 10, 17, 18, 24
foreach ($col in $textCols) {
    Set-ExcelColumn -Worksheet $sheet -Column $col -NumberFormat "@"
}

# 3. Summalarni ajratilgan ko'rinishga o'tkazish
$moneyCols = 28, 29, 31, 32
foreach ($col in $moneyCols) {
    Set-ExcelColumn -Worksheet $sheet -Column $col -NumberFormat "#,##0.00"
}

Close-ExcelPackage $excel

Write-Host "Zo'r! $script:jsonCount ta JSON fayldan barcha detallar '$exportPath' ga saqlandi." -ForegroundColor Green