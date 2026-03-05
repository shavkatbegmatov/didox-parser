Add-Type -AssemblyName System.IO.Compression.FileSystem

# Joriy papkadagi barcha .zip fayllarni topish
$zipFiles = Get-ChildItem -Path . -Filter "*.zip"
$results = @()

if ($zipFiles.Count -eq 0) {
    Write-Host "Ushbu papkada ZIP fayllar topilmadi." -ForegroundColor Yellow
    exit
}

$jsonCount = 0

foreach ($zip in $zipFiles) {
    $archive = [System.IO.Compression.ZipFile]::OpenRead($zip.FullName)
    
    foreach ($entry in $archive.Entries) {
        if ($entry.FullName.EndsWith(".json", [System.StringComparison]::OrdinalIgnoreCase)) {
            $jsonCount++
            
            $stream = $entry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $jsonString = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()

            $data = $jsonString | ConvertFrom-Json

            # Asosiy ma'lumotlarni o'qish (Hech qanday apostroflarsiz toza olinadi)
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
                $results += $row
            }
        }
    }
    $archive.Dispose()
}

if ($jsonCount -eq 0) {
    Write-Host "ZIP arxivlar topildi, lekin ichida JSON fayllar yo'q." -ForegroundColor Yellow
    exit
}

$exportPath = "Fakturalar_hisoboti.xlsx"

# 1-qadam: Ma'lumotlarni Excelga yozamiz va -PassThru orqali faylni xotirada ochiq qoldiramiz
$excel = $results | Export-Excel -Path $exportPath -AutoSize -BoldTopRow -FreezeTopRow -PassThru

# 2-qadam: Excelning 1-varag'ini (Sheet) tanlaymiz
$sheet = $excel.Workbook.Worksheets[1]

# 3-qadam: Ustunlar formatini o'zgartiramiz
# 4-ustun (D) va 6-ustunni (F) to'liq raqamli ko'rinishga ("0" formati) o'tkazamiz
Set-ExcelColumn -Worksheet $sheet -Column 4 -NumberFormat "0"
Set-ExcelColumn -Worksheet $sheet -Column 6 -NumberFormat "0"

# Qo'shimcha: 9, 10 va 11-ustunlardagi summalarni pul formatida chiroyli qilib ajratamiz
Set-ExcelColumn -Worksheet $sheet -Column 9 -NumberFormat "#,##0.00"
Set-ExcelColumn -Worksheet $sheet -Column 10 -NumberFormat "#,##0.00"
Set-ExcelColumn -Worksheet $sheet -Column 11 -NumberFormat "#,##0.00"

# 4-qadam: O'zgarishlarni saqlab, Excel faylni yopamiz
Close-ExcelPackage $excel

Write-Host "Muvaffaqiyatli yakunlandi! $jsonCount ta JSON fayl o'qildi va '$exportPath' Excel fayliga saqlandi." -ForegroundColor Green