# URLs directes extraites de diablo2.io/uniques/
$items = @(
    @{name="Defender's Bile";              url="https://diablo2.io/uniques/defender-s-bile-t1674270.html";              inv="invdfb"},
    @{name="Defender's Fire";              url="https://diablo2.io/uniques/defender-s-fire-t1674269.html";              inv="invdff"},
    @{name="Entropy Locket";               url="https://diablo2.io/uniques/entropy-locket-t1673931.html";               inv="inventl"},
    @{name="Guardian's Light";             url="https://diablo2.io/uniques/guardian-s-light-t1674266.html";             inv="invgdl"},
    @{name="Guardian's Thunder";           url="https://diablo2.io/uniques/guardian-s-thunder-t1674060.html";           inv="invgdt"},
    @{name="Latent Black Cleft";           url="https://diablo2.io/uniques/latent-black-cleft-t1674286.html";           inv="invlbc"},
    @{name="Latent Bone Break";            url="https://diablo2.io/uniques/latent-bone-break-t1674288.html";            inv="invlbb2"},
    @{name="Latent Cold Rupture";          url="https://diablo2.io/uniques/latent-cold-rupture-t1674284.html";          inv="invlcr"},
    @{name="Latent Crack of the Heavens";  url="https://diablo2.io/uniques/latent-crack-of-the-heavens-t1674290.html";  inv="invlch"},
    @{name="Latent Flame Rift";            url="https://diablo2.io/uniques/latent-flame-rift-t1674292.html";            inv="invlfr"},
    @{name="Latent Rotting Fissure";       url="https://diablo2.io/uniques/latent-rotting-fissure-t1674294.html";       inv="invlrf"},
    @{name="Opalvein";                     url="https://diablo2.io/uniques/opalvein-t1673929.html";                     inv="invopal"},
    @{name="Protector's Frost";            url="https://diablo2.io/uniques/protector-s-frost-t1674061.html";            inv="invptf"},
    @{name="Protector's Stone";            url="https://diablo2.io/uniques/protector-s-stone-t1674268.html";            inv="invpts"},
    @{name="Renewed Black Cleft";          url="https://diablo2.io/uniques/renewed-black-cleft-t1674287.html";          inv="invrbc"},
    @{name="Renewed Bone Break";           url="https://diablo2.io/uniques/renewed-bone-break-t1674289.html";           inv="invrbb"},
    @{name="Renewed Cold Rupture";         url="https://diablo2.io/uniques/renewed-cold-rupture-t1674285.html";         inv="invrcr"},
    @{name="Renewed Crack of the Heavens"; url="https://diablo2.io/uniques/renewed-crack-of-the-heavens-t1674291.html"; inv="invrch"},
    @{name="Renewed Flame Rift";           url="https://diablo2.io/uniques/renewed-flame-rift-t1674293.html";           inv="invrfr"},
    @{name="Renewed Rotting Fissure";      url="https://diablo2.io/uniques/renewed-rotting-fissure-t1674295.html";      inv="invrrf"},
    @{name="Skull Collector";              url="https://diablo2.io/uniques/skull-collector-t910.html";                  inv="invskco"},
    @{name="Sling";                        url="https://diablo2.io/uniques/sling-t1673930.html";                        inv="invslng"},
    @{name="Wraithstep";                   url="https://diablo2.io/uniques/wraithstep-t1673927.html";                   inv="invwrst"}
)

foreach ($item in $items) {
    try {
        $r = Invoke-WebRequest -Uri $item.url -UseBasicParsing -TimeoutSec 10
        $graphic = [regex]::Matches($r.Content, 'images/items/([^"]+_graphic[^"]*\.png)') | ForEach-Object { $_.Groups[1].Value } | Select-Object -First 1
        Write-Host "$($item.name) [$($item.inv)] -> $graphic"
    } catch {
        Write-Host "$($item.name) -> FAIL" -ForegroundColor Red
    }
}
