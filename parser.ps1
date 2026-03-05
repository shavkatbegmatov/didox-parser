Add-Type -AssemblyName System.IO.Compression.FileSystem

# Dastlabki o'zgaruvchilarni tayyorlaymiz
$script:results = @()
$script:jsonCount = 0

$zipFiles = Get-ChildItem -Path . -Filter "*.zip"

if ($zipFiles.Count -eq 0) {
    Write-Host "Ushbu papkada ZIP fayllar topilmadi." -ForegroundColor Yellow
    exit
}

# Arxivlarni ichma-ich o'quvchi Maxsus Funksiya (Rekursiya)
function Read-ZipArchive {
    param (
        [System.IO.Stream]$Stream
    )

    try {
        $archive = New-Object System.IO.Compression.ZipArchive($Stream, [System.IO.Compression.ZipArchiveMode]::Read)
        
        foreach ($entry in $archive.Entries) {
            # 1. Agar fayl JSON bo'lsa, uni o'qiymiz
            if ($entry.FullName.EndsWith(".json", [System.StringComparison]::OrdinalIgnoreCase)) {
                $script:jsonCount++
                
                $entryStream = $entry.Open()
                $reader = New-Object System.IO.StreamReader($entryStream)
                $jsonString = $reader.ReadToEnd()
                $reader.Close()
                $entryStream.Close()

                $data = $jsonString | ConvertFrom-Json

                # Ma'lumotlarni ajratish
                $fakturaNo = $data.facturadoc.facturano
                $fakturaDate = $data.facturadoc.facturadate
                $sellerName = $data.seller.name
                $sellerStir = $data.seller.vatregcode 
                $buyerName = $data.buyer.name
                $buyerStir = $data.buyer.vatregcode 

                foreach ($product in $data.productlist.products) {
                    $row = [PSCustomObject]@{
                        "Hujjat Raqami" = $fakturaNo
                        "Hujjat Sanasi" = $fakturaDate
                        "Sotuvchi Tashkilot" = $sellerName
                        "Sotuvchi STIR" = $sellerStir
                        "Xaridor Tashkilot" = $buyerName
                        "Xaridor STIR" = $buyerStir
                        "Xizmat / Mahsulot Nomi" = $product.name
                        "Soni" = $product.count
                        "Yetkazib Berish Narxi (QQSsiz)" = $product.deliverysum
                        "QQS Summasi" = $product.vatsum
                        "Jami Summa (QQS bilan)" = $product.deliverysumwithvat
                    }
                    $script:results += $row
                }
            }
            # 2. Agar fayl ZIP bo'lsa, uni xotiraga yuklab o'zini o'zi qayta chaqiramiz
            elseif ($entry.FullName.EndsWith(".zip", [System.StringComparison]::OrdinalIgnoreCase)) {
                $innerStream = $entry.Open()
                $memStream = New-Object System.IO.MemoryStream
                $innerStream.CopyTo($memStream)
                $innerStream.Close()
                
                $memStream.Position = 0
                
                # Arxiv ichidagi arxivni o'qish (Rekursiya qadami)
                Read-ZipArchive -Stream $memStream
                
                $memStream.Dispose()
            }
        }
        $archive.Dispose()
    }
    catch {
        Write-Host "Arxivni o'qishda xatolik yuz berdi: $_" -ForegroundColor Red
    }
}

Write-Host "Arxivlar tahlil qilinmoqda, kuting..." -ForegroundColor Cyan

# Asosiy papkadagi barcha ZIP fayllarni birma-bir ochib funksiyaga uzatamiz
foreach ($zip in $zipFiles) {
    $fileStream = [System.IO.File]::OpenRead($zip.FullName)
    Read-ZipArchive -Stream $fileStream
    $fileStream.Close()
}

if ($script:jsonCount -eq 0) {
    Write-Host "Hech qanday JSON fayl topilmadi." -ForegroundColor Yellow
    exit
}

$exportPath = "Fakturalar_hisoboti.xlsx"

# Excel faylga yozish va formatlash
$excel = $script:results | Export-Excel -Path $exportPath -AutoSize -BoldTopRow -FreezeTopRow -PassThru
$sheet = $excel.Workbook.Worksheets[1]

# STIR raqamlarini Exceledagi E+ ko'rinishidan qutqarib, "To'liq raqam" formatiga o'tkazish
Set-ExcelColumn -Worksheet $sheet -Column 4 -NumberFormat "0"
Set-ExcelColumn -Worksheet $sheet -Column 6 -NumberFormat "0"

# Summalarni o'qishga oson bo'lishi uchun pul formatiga (bo'shliq va yuzlik bilan) o'tkazish
Set-ExcelColumn -Worksheet $sheet -Column 9 -NumberFormat "#,##0.00"
Set-ExcelColumn -Worksheet $sheet -Column 10 -NumberFormat "#,##0.00"
Set-ExcelColumn -Worksheet $sheet -Column 11 -NumberFormat "#,##0.00"

Close-ExcelPackage $excel

Write-Host "Muvaffaqiyatli yakunlandi! Jami $script:jsonCount ta JSON fayl (ichma-ich arxivlardan) o'qildi va '$exportPath' ga saqlandi." -ForegroundColor Green