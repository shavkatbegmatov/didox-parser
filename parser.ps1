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
    # ZIP arxivni o'qish uchun ochamiz
    $archive = [System.IO.Compression.ZipFile]::OpenRead($zip.FullName)
    
    foreach ($entry in $archive.Entries) {
        # Agar fayl .json bo'lsa
        if ($entry.FullName.EndsWith(".json", [System.StringComparison]::OrdinalIgnoreCase)) {
            $jsonCount++
            
            # JSON faylni arxivdan o'qish
            $stream = $entry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $jsonString = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()

            # JSON matnni aylantirish (xatolik bermasligi uchun Depth olib tashlandi)
            $data = $jsonString | ConvertFrom-Json

            # Asosiy ma'lumotlarni olish
            $fakturaNo = $data.facturadoc.facturano
            $fakturaDate = $data.facturadoc.facturadate
            $sellerName = $data.seller.name
            $sellerStir = $data.seller.vatregcode
            $buyerName = $data.buyer.name
            $buyerStir = $data.buyer.vatregcode

            # Mahsulotlarni jadval qatorlariga yig'ish
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

# Natijani to'g'ridan-to'g'ri haqiqiy Excel (.xlsx) faylga saqlash
$exportPath = "Fakturalar_hisoboti.xlsx"
$results | Export-Excel -Path $exportPath -AutoSize -BoldTopRow -FreezeTopRow

Write-Host "Muvaffaqiyatli yakunlandi! $jsonCount ta JSON fayl o'qildi va '$exportPath' Excel fayliga saqlandi." -ForegroundColor Green