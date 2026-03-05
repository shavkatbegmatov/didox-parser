# ZIP fayllar bilan ishlash uchun .NET kutubxonasini chaqiramiz
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
            
            # JSON faylni arxivdan to'g'ridan-to'g'ri xotiraga o'qiymiz
            $stream = $entry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $jsonString = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()

            # JSON matnni PowerShell obyektiga aylantirish
            $data = $jsonString | ConvertFrom-Json

            # Asosiy ma'lumotlarni olish
            $fakturaNo = $data.facturadoc.facturano
            $fakturaDate = $data.facturadoc.facturadate
            $sellerName = $data.seller.name
            $sellerStir = $data.seller.vatregcode
            $buyerName = $data.buyer.name
            $buyerStir = $data.buyer.vatregcode

            # Har bir xizmat/mahsulot uchun qator yaratish
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
    # Arxivni yopish
    $archive.Dispose()
}

if ($jsonCount -eq 0) {
    Write-Host "ZIP arxivlar topildi, lekin ichida JSON fayllar yo'q." -ForegroundColor Yellow
    exit
}

# Natijani CSV faylga saqlash (Excel bemalol ocha oladi)
$exportPath = "Fakturalar_hisoboti.csv"
$results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

Write-Host "Muvaffaqiyatli yakunlandi! $jsonCount ta JSON fayl o'qildi va '$exportPath' fayliga saqlandi." -ForegroundColor Green