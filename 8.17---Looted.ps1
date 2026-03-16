# ================================================================
# Looted.ps1 - Discord Item Logger for Diablo 2 Resurrected
# VERSION FINALE COMPLÈTE - FULL SUPPORT
# ================================================================
# CORRECTIONS: v
# - Détection RARE vs UNIQUE corrigée (plus de rares dans unique)
# - Fuzzy matching strict pour éviter faux positifs
# - Fonction Close-D2RInstance NUCLEAR (5 méthodes progressives)
# - PID mapping automatique depuis les logs
# - Reset des deaths consécutives si retour en ville
# - [NOUVEAU] Reset automatique "Deaths today" après 11h d'inactivité
# - [NOUVEAU] Détection WHITE améliorée (Socketed, All Res, Enh Def, etc.)
# - [NOUVEAU] Double reset "Consecutive Deaths" renforcé après kill
# ================================================================

$ErrorActionPreference = "SilentlyContinue"
Set-Location $PSScriptRoot

# Paths
$configPath = Join-Path $PSScriptRoot "config.json"
$dataPath = Join-Path $PSScriptRoot "data"
$d2dataPath = Join-Path $dataPath "d2data"


# Global data
$global:gfx = @{}
$global:itemData = @{}
$global:miscItems = @{}
$global:uniqueItems = @{}
$global:setItems = @{}
$global:runewords = @{}
$global:prefixRare = @{}
$global:suffixRare = @{}
$global:prefixMagic = @{}
$global:suffixMagic = @{}
$global:charMapping = @{}
$global:nameToTagId = @{}
$global:deathCount = @{}
$global:reportedDeaths = @{}
$global:sceneMap = @{}
$global:townScenes = @()
$global:pidMapping = @{}
$global:consecutiveDeaths = @{}

# [NOUVEAU] Tracking des derniers timestamps de mort pour reset automatique
$global:lastDeathTimestamps = @{}

# ================================================================
# UNIQUE ITEM NAME MAPPING
# ================================================================
# Mappage des noms d'items qui diffèrent entre le jeu et le JSON
# Format: "Nom dans le jeu" = "Nom dans uniqueitems.json"
$global:uniqueNameMapping = @{
    "Wisp Projector" = "Wisp"
    "Verdungo's Hearty Cord" = "Verdugo's Hearty Cord"
    "Cerebus' Bite" = "Cerebus"
    "Bul-Kathos' Wedding Band" = "Bul Katho's Wedding Band"

    # Ajoutez ici d'autres mappings si nécessaire
}

# ================================================================
# UNIVERSAL DATE PARSING
# ================================================================

function Parse-Timestamp {
    param(
        [string]$Timestamp,
        [switch]$IsUTC  # Si $true, le timestamp est en UTC et sera converti en local
    )

    if (-not $Timestamp) { return $null }

    # Liste des formats possibles
    $formats = @(
        "dd-MM-yyyy HH:mm:ss",      # Format Looted: 30-10-2025 23:24:29
        "yyyy-MM-dd HH:mm:ss",      # Format D2R logs: 2025-10-30 23:24:29
        "dd/MM/yyyy HH:mm:ss",      # Format alternatif: 30/10/2025 23:24:29
        "yyyy/MM/dd HH:mm:ss",      # Format alternatif: 2025/10/30 23:24:29
        "MM-dd-yyyy HH:mm:ss",      # Format US: 10-30-2025 23:24:29
        "MM/dd/yyyy HH:mm:ss"       # Format US: 10/30/2025 23:24:29
    )

    $dt = $null

    # Essayer chaque format
    foreach ($format in $formats) {
        try {
            $dt = [DateTime]::ParseExact($Timestamp, $format, $null)
            break
        }
        catch {
            continue
        }
    }

    # Si aucun format ne fonctionne, essayer le parse automatique
    if (-not $dt) {
        try {
            $dt = [DateTime]::Parse($Timestamp)
        }
        catch {
            return $null
        }
    }

    # Si le timestamp est en UTC, le convertir en heure locale
    if ($IsUTC -and $dt) {
        $dt = [DateTime]::SpecifyKind($dt, [DateTimeKind]::Utc)
        $dt = $dt.ToLocalTime()
    }

    return $dt
}


# ================================================================
# DATA LOADING
# ================================================================

function Load-D2Data {
    Write-Host "Loading D2R data files..." -ForegroundColor Cyan
    
    if (Test-Path "$d2dataPath\misc.json") {
        try {
            $miscJson = Get-Content "$d2dataPath\misc.json" -Raw | ConvertFrom-Json
            $miscJson.PSObject.Properties | ForEach-Object {
                $code = $_.Name
                $item = $_.Value
                $global:miscItems[$code] = $item
                if ($item.name -and $item.invcode) {
                    $global:miscItems[$item.name] = $item
                    $normalizedName = $item.name -replace "[''`]", "" -replace "\s+", " "
                    if ($normalizedName -ne $item.name) {
                        $global:miscItems[$normalizedName] = $item
                    }
                }
            }
        }
        catch { }
    }
    
    if (Test-Path "$d2dataPath\item.json") {
        try {
            $itemsJson = Get-Content "$d2dataPath\item.json" -Raw | ConvertFrom-Json
            $itemsJson.PSObject.Properties | ForEach-Object {
                $code = $_.Name
                $item = $_.Value
                $global:itemData[$code] = $item
                if ($item.name -and $item.invfile) {
                    $global:itemData[$item.name] = $item
                    $normalizedName = $item.name -replace "[''`]", "" -replace "\s+", " "
                    if ($normalizedName -ne $item.name) {
                        $global:itemData[$normalizedName] = $item
                    }
                }
            }
        }
        catch { }
    }

    if (Test-Path "$dataPath\scenes.json") {
        try {
            $scenesData = Get-Content "$dataPath\scenes.json" -Raw | ConvertFrom-Json
            $scenesData.scenes.PSObject.Properties | ForEach-Object {
                $global:sceneMap[$_.Name] = $_.Value
            }
            $global:townScenes = $scenesData.towns
        }
        catch {
        }
    }
    
    if (Test-Path "$dataPath\gfx.csv") {
        $gfxData = Import-Csv "$dataPath\gfx.csv" -Header name, invfile
        foreach ($row in $gfxData) {
            if ($row.name) {
                $global:gfx[$row.name] = $row.invfile
            }
        }
    }
    
    # Charger uniquenames.json au lieu de uniqueitems.json
    if (Test-Path "$dataPath\uniquenames.json") {
        try {
            $uniqueNamesData = Get-Content "$dataPath\uniquenames.json" -Raw | ConvertFrom-Json
            
            foreach ($uniqueName in $uniqueNamesData.uniqueNames) {
                $global:uniqueItems[$uniqueName] = $true
                $normalized = $uniqueName -replace "[''`]", "" -replace "\s+", " "
                if ($normalized -ne $uniqueName) {
                    $global:uniqueItems[$normalized] = $true
                }
            }
            
            Write-Host "  ✓ Loaded $($uniqueNamesData.totalItems) unique items" -ForegroundColor Green
        }
        catch {
            Write-Host "  ✗ Error loading uniquenames.json: $_" -ForegroundColor Red
        }
    }
    elseif (Test-Path "$d2dataPath\uniqueitems.json") {
        $uniqueData = Get-Content "$d2dataPath\uniqueitems.json" -Raw | ConvertFrom-Json
        $uniqueData.PSObject.Properties | ForEach-Object {
            $item = $_.Value
            
            if ($item.index) {
                $global:uniqueItems[$item.index] = $true
                $normalized = $item.index -replace "[''`]", "" -replace "\s+", " "
                $global:uniqueItems[$normalized] = $true
            }
            
            if ($item.name -and $item.name -ne $item.index) {
                $global:uniqueItems[$item.name] = $true
                $normalized = $item.name -replace "[''`]", "" -replace "\s+", " "
                $global:uniqueItems[$normalized] = $true
            }
        }
    }
    
    if (Test-Path "$d2dataPath\setitems.json") {
        $setData = Get-Content "$d2dataPath\setitems.json" -Raw | ConvertFrom-Json
        $setData.PSObject.Properties | ForEach-Object {
            $item = $_.Value
            
            if ($item.index) {
                $global:setItems[$item.index] = $item
                $normalized = $item.index -replace "[''`]", "" -replace "\s+", " "
                $global:setItems[$normalized] = $item
            }
            
            if ($item.name -and $item.name -ne $item.index) {
                $global:setItems[$item.name] = $item
                $normalized = $item.name -replace "[''`]", "" -replace "\s+", " "
                $global:setItems[$normalized] = $item
            }
        }
    }
    
    if (Test-Path "$d2dataPath\runes.json") {
        $runeData = Get-Content "$d2dataPath\runes.json" -Raw | ConvertFrom-Json
        $runeData.PSObject.Properties | ForEach-Object {
            $runeName = $_.Name
            $runeDetails = $_.Value
            $global:runewords[$runeName] = $runeDetails
            $normalized = $runeName -replace "[''`]", "" -replace "\s+", " "
            if ($normalized -ne $runeName) {
                $global:runewords[$normalized] = $runeDetails
            }
        }
    }
    
    if (Test-Path "$dataPath\prefix_rare.csv") {
        $prefixData = Import-Csv "$dataPath\prefix_rare.csv" -Header chinese, english
        foreach ($row in $prefixData) {
            if ($row.english) {
                $global:prefixRare[$row.english.ToLower()] = $true
            }
        }
    }
    
    if (Test-Path "$dataPath\suffix_rare.csv") {
        $suffixData = Import-Csv "$dataPath\suffix_rare.csv" -Header chinese, english
        foreach ($row in $suffixData) {
            if ($row.english) {
                $global:suffixRare[$row.english.ToLower()] = $true
            }
        }
    }
    
    if (Test-Path "$dataPath\prefix_magic.csv") {
        $prefixData = Import-Csv "$dataPath\prefix_magic.csv" -Header chinese, english
        foreach ($row in $prefixData) {
            if ($row.english) {
                $global:prefixMagic[$row.english.ToLower()] = $true
            }
        }
    }
    
    if (Test-Path "$dataPath\suffix_magic.csv") {
        $suffixData = Import-Csv "$dataPath\suffix_magic.csv" -Header chinese, english
        foreach ($row in $suffixData) {
            if ($row.english) {
                $global:suffixMagic[$row.english.ToLower()] = $true
            }
        }
    }
    
}

function Load-CharacterMapping {
    param([string]$SettingsPath)
    
    if (-not (Test-Path $SettingsPath)) {
        return
    }
    
    try {
        $content = Get-Content $SettingsPath -Raw
        $isJson = $content.TrimStart().StartsWith("{")
        
        if ($isJson) {
            $config = $content | ConvertFrom-Json
            if ($config.Bots) {
                foreach ($bot in $config.Bots) {
                    $tagId = $bot.TagId
                    $charName = $bot.CharacterName
                    if ($tagId -and $charName) {
                        $global:charMapping[$tagId] = $charName
                        $global:nameToTagId[$charName.ToLower()] = $tagId
                    }
                }
            }
        }
        else {
            [xml]$settingsXml = $content
            $charElements = $settingsXml.SelectNodes("//CharacterSettings/Char")
            foreach ($charElement in $charElements) {
                $tagId = $charElement.GetAttribute("TagId")
                $charName = $charElement.GetAttribute("Name")
                if ($tagId -and $charName) {
                    $global:charMapping[$tagId] = $charName
                    $global:nameToTagId[$charName.ToLower()] = $tagId
                }
            }
        }
    }
    catch { }
}

function Load-Config {
    param([string]$Path)
    
    if (Test-Path $Path) {
        try {
            $configJson = Get-Content $Path -Raw | ConvertFrom-Json
            
            $config = @{}
            $configJson.PSObject.Properties | ForEach-Object {
                if ($_.Value -is [System.Management.Automation.PSCustomObject]) {
                    $nested = @{}
                    $_.Value.PSObject.Properties | ForEach-Object {
                        $nested[$_.Name] = $_.Value
                    }
                    $config[$_.Name] = $nested
                }
                else {
                    $config[$_.Name] = $_.Value
                }
            }
            
            Write-Host "Configuration loaded" -ForegroundColor Green
            return $config
        }
        catch {
            Write-Host "Error loading config: $_" -ForegroundColor Red
            return $null
        }
    }
    return $null
}

# ================================================================
# ITEM PARSING
# ================================================================

function Get-ItemCount {
    param([string]$Path)
    
    try {
        $content = Get-Content $Path -Raw -Encoding Default
        $pattern = '(?m)(?<=^)([\s\S]+?)(?=\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2}|\Z)'
        $matches = [regex]::Matches($content, $pattern)
        return $matches.Count
    }
    catch {
        return 0
    }
}

function Get-ItemFromLog {
    param(
        [string]$Path,
        [int]$Index,
        [string]$PlayerName
    )
    
    try {
        $content = Get-Content $Path -Raw -Encoding Default
        $pattern = '(?m)(?<=^)([\s\S]+?)(?=\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2}|\Z)'
        $matches = [regex]::Matches($content, $pattern)
        
        if ($Index -lt $matches.Count) {
            $lines = $matches[$Index].Value.Trim() -split "`n"
            
            if ($lines.Count -ge 2) {
                $timestamp = $lines[0].Trim().TrimEnd('.:')
                $itemName = $lines[1].Trim()
                # Supprime le préfixe de formatage D2R: U+E07E (rectangle ?) + code couleur (lettre)
                # Ex: "[U+E07E]JIth Rune" → "Ith Rune", "[U+E07E]LTwisted Essence..." → "Twisted Essence..."
                $itemName = $itemName -replace '[\uE000-\uF8FF][A-Za-z]', ''
                
                # Certains items (Keys, Essences) n'ont que 2 lignes
                $itemType = ""
                $stats = ""
                
                if ($lines.Count -ge 3) {
                    $itemType = $lines[2].Trim()
                    
                    if ($lines.Count -gt 3) {
                        $stats = ($lines[3..($lines.Count - 1)] | ForEach-Object { $_.Trim() }) -join "`n"
                    }
                }
                
                return @{
                    Name = $itemName
                    Type = $itemType
                    Stats = $stats
                    Timestamp = $timestamp
                    FoundBy = $PlayerName
                    AllLines = $lines
                }
            }
        }
    }
    catch {
        Write-Host "Error parsing item: $_" -ForegroundColor Red
    }
    
    return $null
}

function Get-GFX {
    param(
        [string]$ItemName,
        [string]$ItemType
    )
    
    # Priority -1: Check misc.json (uses invcode)
    if ($global:miscItems.ContainsKey($ItemName)) {
        $miscInfo = $global:miscItems[$ItemName]
        if ($miscInfo.invcode) {
            return $miscInfo.invcode
        }
    }
    
    $normalizedName = $ItemName -replace "[''`]", "" -replace "\s+", " "
    if ($normalizedName -ne $ItemName -and $global:miscItems.ContainsKey($normalizedName)) {
        $miscInfo = $global:miscItems[$normalizedName]
        if ($miscInfo.invcode) {
            return $miscInfo.invcode
        }
    }
    
    # Priority 0: Check item.json (uses invfile)
    if ($global:itemData.ContainsKey($ItemName)) {
        $itemInfo = $global:itemData[$ItemName]
        if ($itemInfo.invfile) {
            return $itemInfo.invfile
        }
    }
    
    if ($normalizedName -ne $ItemName -and $global:itemData.ContainsKey($normalizedName)) {
        $itemInfo = $global:itemData[$normalizedName]
        if ($itemInfo.invfile) {
            return $itemInfo.invfile
        }
    }
    # Priority 1: Check charms and jewels
    if ($ItemName -match 'Small Charm' -and $global:gfx.ContainsKey('Small Charm')) {
        return $global:gfx['Small Charm']
    }
    if ($ItemName -match 'Large Charm' -and $global:gfx.ContainsKey('Large Charm')) {
        return $global:gfx['Large Charm']
    }
    if ($ItemName -match 'Grand Charm' -and $global:gfx.ContainsKey('Grand Charm')) {
        return $global:gfx['Grand Charm']
    }
    # Match ONLY "Jewel" but NOT "Jeweler's" (magic prefix)
    if ($ItemName -match '\bJewel\b' -and $ItemName -notmatch '\bJeweler') {
        if ($global:gfx.ContainsKey('Jewel')) {
            return $global:gfx['Jewel']
        }
    }
    
    
    # Priority 2: Exact item name match
    if ($global:gfx.ContainsKey($ItemName)) {
        return $global:gfx[$ItemName]
    }
    
    # Priority 3: Item type match
    if ($global:gfx.ContainsKey($ItemType)) {
        return $global:gfx[$ItemType]
    }
    
    # Priority 3.5: Smart extraction for MAGIC items only
    # Magic items format: "Prefix BaseType of Suffix" or "Prefix BaseType" or "BaseType of Suffix"
    # Examples: "Gaean Amulet of the Apprentice" → extract "Amulet"
    #           "Cunning Circlet of the Magus" → extract "Circlet"
    if ($ItemName -match '\s+of\s+') {
        # Item has " of " pattern, likely magic with suffix
        $parts = $ItemName -split '\s+of\s+', 2
        $beforeOf = $parts[0].Trim()
        
        # Extract last word before " of " as base type
        $wordsBeforeOf = $beforeOf -split '\s+'
        if ($wordsBeforeOf.Count -ge 2) {
            $baseType = $wordsBeforeOf[-1]
            if ($global:gfx.ContainsKey($baseType)) {
                return $global:gfx[$baseType]
            }
        }
    }
    
    # Priority 4: Progressive word fallback (for rare/magic items)
    # Tries multi-word combinations first to avoid false matches
    # Example: "Gemmed Ring Mail" -> tries "Ring Mail" (armor) before "Mail" or "Ring" (accessory)
    # Example: "Crimson Mesh" -> try "Mesh", "Venomous Battle Axe" -> try "Battle Axe" then "Axe"
    $words = $ItemName -split '\s+'
    for ($i = 1; $i -lt $words.Count; $i++) {
        $baseType = ($words[$i..($words.Count-1)] -join ' ')
        if ($global:gfx.ContainsKey($baseType)) {
            return $global:gfx[$baseType]
        }
    }

    # Priority 5: Single keyword fallback for accessories/equipment types (handles plurals)
    # Runs AFTER progressive fallback to avoid "Ring" matching before "Ring Mail"
    foreach ($word in $words) {
        if ($word -match '^(Amulet|Ring|Circlet|Coronet|Diadem|Tiara|Cap|Helm|Mask|Shield|Armor|Belt|Gloves|Boots|Axe|Sword|Mace|Spear|Bow|Staff|Wand|Scepter|Polearm|Javelin|Crossbow|Dagger|Club)s?$') {
            $cleanWord = $word -replace 's$', ''  # Remove plural 's'
            if ($global:gfx.ContainsKey($cleanWord)) {
                return $global:gfx[$cleanWord]
            }
            if ($global:gfx.ContainsKey($word)) {
                return $global:gfx[$word]
            }
        }
    }
    
    # Default fallback
    return "invgbi"
}

function Get-ImageUrl {
    param([string]$InvFile)
    
    $baseUrl = "https://raw.githubusercontent.com/potlo164/Looted-gid-Potlo/master/assets/gfx_hd_low"
    return "$baseUrl/$InvFile/21.png?v=7"
}

function Get-SceneName {
    param([string]$SceneId)
    
    if ($global:sceneMap.ContainsKey($SceneId)) {
        return $global:sceneMap[$SceneId]
    }
    
    return $null
}

function Test-IsRareItem {
    param(
        [string]$ItemName,
        [string]$ItemStats
    )
    
    $words = $ItemName -split '\s+'
    
    if ($words.Count -lt 2) {
        return $false
    }
    
    $firstWord = $words[0].ToLower()
    $lastWord = $words[$words.Count - 1].ToLower()
    
    $hasRarePrefix = $global:prefixRare.ContainsKey($firstWord)
    $hasRareSuffix = $global:suffixRare.ContainsKey($lastWord)
    
    if ($hasRarePrefix -or $hasRareSuffix) {
        return $true
    }
    
    return $false
}

function Test-IsMagicItem {
    param([string]$ItemName)
    
    $words = $ItemName -split '\s+'
    
    if ($words.Count -lt 2) {
        return $false
    }
    
    $firstWord = $words[0].ToLower()
    $lastWord = $words[$words.Count - 1].ToLower()
    
    $hasMagicPrefix = $global:prefixMagic.ContainsKey($firstWord)
    $hasMagicSuffix = $global:suffixMagic.ContainsKey($lastWord)
    
    if ($hasMagicPrefix -or $hasMagicSuffix) {
        return $true
    }
    
    return $false
}


function Get-FuzzyMatchScore {
    param(
        [string]$String1,
        [string]$String2
    )
    
    # Remove ALL spaces AND convert to lowercase for comparison
    # Handles "Gore Rider" vs "Gorerider" and case differences
    $cleanString1 = ($String1 -replace '\s+', '').ToLower()
    $cleanString2 = ($String2 -replace '\s+', '').ToLower()
    
    $len1 = $cleanString1.Length
    $len2 = $cleanString2.Length
    
    $lengthDiff = [Math]::Abs($len1 - $len2)
    if ($lengthDiff -gt 3) {
        return 0
    }
    
    $differences = 0
    $minLength = [Math]::Min($len1, $len2)
    
    for ($i = 0; $i -lt $minLength; $i++) {
        if ($cleanString1[$i] -ne $cleanString2[$i]) {
            $differences++
        }
    }
    
    $differences += $lengthDiff
    
    if ($differences -gt 3) {
        return 0
    }
    
    return 100 - ($differences * 10)
}

# ================================================================
# [CORRECTION] Fonction pour vérifier si un item avec prefix/suffix 
# magic/rare est en fait un WHITE avec propriétés autochtones
# ================================================================
function Test-IsActuallyWhiteItem {
    param(
        [string]$ItemStats,
        [string]$ItemType,
        [string]$ItemName
    )
    
    # Rings et Amulets ne peuvent jamais être WHITE
    if ($ItemName -match '\b(Ring|Amulet)\b' -or $ItemType -match '\b(Ring|Amulet)\b') {
        return $false
    }
    
    if ([string]::IsNullOrWhiteSpace($ItemStats)) {
        return $false
    }
    
    # Liste complète des propriétés autochtones WHITE (incluant socketed, all res, etc.)
    $basePropertyKeywords = @(
        # Propriétés de base classiques
        'Defense:', 'Damage:', 'Durability:', 'Required', 'Class:', 'Attack Speed:',
        'One-Hand', 'Two-Hand', 'Throw Damage:', 'Quantity:', 'Stack Size:',
        'Smite Damage:', 'Kick Damage:', 'Can be Inserted', 'Chance to Block:',
        
        # Restrictions de classe
        'Assassin Only', 'Paladin Only', 'Necromancer Only', 'Sorceress Only',
        'Amazon Only', 'Druid Only', 'Barbarian Only',
        
        # Types d'armes/armures
        'Sword', 'Axe', 'Bow', 'Crossbow', 'Dagger', 'Mace', 'Spear',
        'Staff', 'Wand', 'Club', 'Hammer', 'Scepter', 'Polearm', 'Javelin',
        
        # [CORRECTION] Propriétés autochtones des WHITE items supérieurs/socketés
        'All Resistances', 'All Resistance',        # Paladin shields
        'Enhanced Defense',                          # Superior defense
        'Increased Maximum Durability',              # Superior durability
        'Enhanced Damage',                           # Superior damage
        'Increased Attack Speed',                    # Superior IAS
        'Replenish', 'Repairs'                       # Réparation auto
    )
    
    $statLines = $ItemStats -split "`n"
    
    foreach ($statLine in $statLines) {
        $cleanLine = $statLine.Trim()
        if ([string]::IsNullOrWhiteSpace($cleanLine)) { continue }
        
        $isBaseProperty = $false
        
        # Vérifier les keywords
        foreach ($keyword in $basePropertyKeywords) {
            if ($cleanLine -like "*$keyword*") {
                $isBaseProperty = $true
                break
            }
        }
        
        # Patterns supplémentaires pour propriétés WHITE
        if (-not $isBaseProperty) {
            # Socketed (2), Socketed (3), Socketed (4), etc.
            if ($cleanLine -match '^Socketed \(\d+\)$') {
                $isBaseProperty = $true
            }
            # Juste un nombre seul : "3", "45", etc.
            elseif ($cleanLine -match '^\d+$') {
                $isBaseProperty = $true
            }
            # Juste (nombre) : "(3)", "(4)", etc.
            elseif ($cleanLine -match '^\(\d+\)$') {
                $isBaseProperty = $true
            }
            # +nombre : "+45", "+15", etc.
            elseif ($cleanLine -match '^\+\d+$') {
                $isBaseProperty = $true
            }
            # Durability format : "34 of 40"
            elseif ($cleanLine -match '^\d+ of \d+$') {
                $isBaseProperty = $true
            }
            # Restrictions de classe : "(Paladin Only)"
            elseif ($cleanLine -match '^\([A-Za-z ]+Only\)$') {
                $isBaseProperty = $true
            }
        }
        
        # Si on trouve UNE stat qui n'est PAS une propriété de base → c'est pas WHITE
        if (-not $isBaseProperty) {
            return $false
        }
    }
    
    # Toutes les stats sont des propriétés de base → c'est WHITE
    return $true
}

function Get-ItemQuality {
    param(
        [string]$ItemName,
        [string]$ItemType,
        [string]$ItemStats,
        [array]$AllLines
    )
    
    $isUnidentified = $false
    foreach ($line in $AllLines) {
        if ($line -match '^Unidentified\s*$') {
            $isUnidentified = $true
            break
        }
    }
    
    # PRIORITÉ 1: Charms et Jewels
    if ($ItemName -match '(Small|Large|Grand) Charm' -or $ItemName -match '\bJewel\b') {
        $charmType = $null
        if ($ItemName -match 'Small Charm') { $charmType = "small_charm" }
        elseif ($ItemName -match 'Large Charm') { $charmType = "large_charm" }
        elseif ($ItemName -match 'Grand Charm') { $charmType = "grand_charm" }
        elseif ($ItemName -match '\bJewel\b') { $charmType = "jewel" }
        
        if ($global:uniqueItems.ContainsKey($ItemName)) {
            return "${charmType}_unique"
        }
        
        if ($charmType -eq "jewel") {
            if (Test-IsRareItem -ItemName $ItemName -ItemStats $ItemStats) {
                return "${charmType}_rare"
            }
        }
        
        return "${charmType}_magic"
    }
    
    # PRIORITÉ 2: Unidentified
    if ($isUnidentified) {
        return "unique"
    }
    
    # PRIORITÉ 3: Runes
    if ($ItemName -match '^(Pul|Um|Mal|Ist|Gul|Vex|Ohm|Lo|Sur|Ber|Jah|Cham|Zod) Rune$') {
        return "high_rune"
    }
    
    if ($ItemName -match '^(El|Eld|Tir|Nef|Eth|Ith|Tal|Ral|Ort|Thul|Amn|Sol|Shael|Dol|Hel|Io|Lum|Ko|Fal|Lem) Rune$') {
        return "rune"
    }
    
    # PRIORITÉ 3.5: Misc items
    $normalizedName = $ItemName -replace "[''`]", "" -replace "\s+", " "
    if ($global:miscItems.ContainsKey($ItemName) -or $global:miscItems.ContainsKey($normalizedName)) {
        return "misc"
    }
    
    # PRIORITÉ 3.6: Fallback hardcodé pour Keys et Essences
    if ($ItemName -match '^Key of (Terror|Hate|Destruction)$') {
        return "uber_keys"
    }
    if ($ItemName -match '^(Twisted Essence of Suffering|Charged Essence of Hatred|Burning Essence of Terror|Festering Essence of Destruction|Token of Absolution)$') {
        return "essences"
    }
    
    # PRIORITÉ 4: Unique - Exact match (AVANT WHITE!)
    $normalizedItemName = $ItemName -replace "[''`]", "" -replace "\s+", " "
    
    # Vérifier d'abord dans la table de mapping des noms
    $mappedName = $ItemName
    if ($global:uniqueNameMapping.ContainsKey($ItemName)) {
        $mappedName = $global:uniqueNameMapping[$ItemName]
    }
    elseif ($global:uniqueNameMapping.ContainsKey($normalizedItemName)) {
        $mappedName = $global:uniqueNameMapping[$normalizedItemName]
    }
    
    # Normaliser aussi le nom mappé pour gérer les apostrophes
    $normalizedMappedName = $mappedName -replace "[''`]", "" -replace "\s+", " "
    
    # Vérifier avec le nom mappé ET le nom original
    if ($global:uniqueItems.ContainsKey($ItemName) -or 
        $global:uniqueItems.ContainsKey($normalizedItemName) -or
        $global:uniqueItems.ContainsKey($mappedName) -or
        $global:uniqueItems.ContainsKey($normalizedMappedName)) {
        return "unique"
    }
    
    # PRIORITÉ 5: Set - Exact match (AVANT WHITE!)
    if ($global:setItems.ContainsKey($ItemName) -or $global:setItems.ContainsKey($normalizedItemName)) {
        return "set"
    }
    
    $knownSetPrefixes = @(
        "Tal Rasha", "Griswold", "Trang-Oul", "Immortal King", "M'avina", "Natalya", 
        "Aldur", "Tancred", "Death's Disguise", "Sander's Folly", "Infernal Tools",
        "Berserker's Arsenal", "Angelic Raiment", "Arctic Gear", "Arcanna's Tricks",
        "Vidala's Rig", "Milabrega's Regalia", "Cathan's Traps", "Iratha's Finery",
        "Sigon's Complete Steel", "Hsarus' Defense", "Cow King's Leathers", "Sazabi",
        "Bul-Kathos", "Heaven's Brethren", "Orphan's Call", "The Disciple", "Naj's Ancient Vestige"
    )
    
    foreach ($prefix in $knownSetPrefixes) {
        if ($ItemName -like "$prefix*") {
            return "set"
        }
    }
    
    # PRIORITÉ 6: RARE - avec vérification WHITE override
    if (Test-IsRareItem -ItemName $ItemName -ItemStats $ItemStats) {
        $normalizedName = $ItemName -replace "[''`]", "" -replace "\s+", " "
        
        if ($global:uniqueItems.ContainsKey($ItemName) -or $global:uniqueItems.ContainsKey($normalizedName)) {
            # C'est un unique, pas un rare - continue
        }
        else {
            # [CORRECTION] Vérifier si c'est en fait un WHITE avec prefix/suffix rare mais stats de base
            $isActuallyWhite = Test-IsActuallyWhiteItem -ItemStats $ItemStats -ItemType $ItemType -ItemName $ItemName
            if ($isActuallyWhite) {
                return "white"
            }
            return "rare"
        }
    }
    

    # PRIORITÉ 7: MAGIC - avec vérification WHITE override
    if (Test-IsMagicItem -ItemName $ItemName) {
        # [CORRECTION] Vérifier si c'est en fait un WHITE avec prefix/suffix magic mais stats de base
        $isActuallyWhite = Test-IsActuallyWhiteItem -ItemStats $ItemStats -ItemType $ItemType -ItemName $ItemName
        if ($isActuallyWhite) {
            return "white"
        }
        return "magic"
    }
    
    # Les anneaux et amulettes ne peuvent jamais être WHITE dans D2R
    # Donc si on arrive ici et que c'est un ring/amulet, c'est au minimum magic
    if ($ItemName -match '\b(Ring|Amulet)\b' -or $ItemType -match '\b(Ring|Amulet)\b') {
        return "magic"
    }
    
    # PRIORITÉ 8: White items (cas standard sans prefix/suffix magic/rare)
    # La détection des WHITE avec prefix/suffix est gérée dans PRIORITÉ 6 et 7
    $isRingOrAmulet = ($ItemName -match '\b(Ring|Amulet)\b' -or $ItemType -match '\b(Ring|Amulet)\b')
    
    if (-not $isRingOrAmulet) {
        $hasOnlyBaseProperties = $true
        
        $basePropertyKeywords = @(
            'Defense:', 'Damage:', 'Durability:', 'Required', 'Class:', 'Attack Speed:',
            'One-Hand', 'Two-Hand', 'Throw Damage:', 'Quantity:', 'Stack Size:',
            'Smite Damage:', 'Kick Damage:', 'Can be Inserted', 'Chance to Block:',
            'Assassin Only', 'Paladin Only', 'Necromancer Only', 'Sorceress Only',
            'Amazon Only', 'Druid Only', 'Barbarian Only',
            'Sword', 'Axe', 'Bow', 'Crossbow', 'Dagger', 'Mace', 'Spear',
            'Staff', 'Wand', 'Club', 'Hammer', 'Scepter', 'Polearm', 'Javelin'
        )
        
        if ($ItemStats) {
            $statLines = $ItemStats -split "`n"
            foreach ($statLine in $statLines) {
                $cleanLine = $statLine.Trim()
                if ([string]::IsNullOrWhiteSpace($cleanLine)) { continue }
                
                $isBaseProperty = $false
                foreach ($keyword in $basePropertyKeywords) {
                    if ($cleanLine -like "*$keyword*") {
                        $isBaseProperty = $true
                        break
                    }
                }
                
                if (-not $isBaseProperty) {
                    $hasOnlyBaseProperties = $false
                    break
                }
            }
        }
        
        if ($hasOnlyBaseProperties -and -not [string]::IsNullOrWhiteSpace($ItemStats)) {
            return "white"
        }
    }
    
    # PRIORITÉ 9: Unique - Fuzzy match (APRÈS WHITE!)
    if ($ItemName.Length -ge 5) {
        $bestScore = 0
        $bestMatch = $null
        
        foreach ($uniqueItem in $global:uniqueItems.Values) {
            $uniqueName = if ($uniqueItem.index) { $uniqueItem.index } else { "" }
            if (-not $uniqueName -or $uniqueName.Length -lt 5) { continue }
            
            $normalizedUnique = $uniqueName -replace "[''`]", "" -replace "\s+", " "
            
            $score = Get-FuzzyMatchScore -String1 $normalizedItemName -String2 $normalizedUnique
            
            if ($score -gt $bestScore) {
                $bestScore = $score
                $bestMatch = "unique"
            }
        }
        
        if ($bestScore -ge 70) {
            return $bestMatch
        }
    }
    

    # PRIORITÉ 10: Set - Fuzzy match (APRÈS WHITE!)
    if ($ItemName.Length -ge 5) {
        foreach ($setItem in $global:setItems.Values) {
            $setName = if ($setItem.index) { $setItem.index } else { "" }
            if (-not $setName) { continue }
            
            $normalizedSetName = $setName -replace "[''`]", "" -replace "\s+", " "
            
            if ($normalizedItemName -like "$normalizedSetName*") {
                $remainingText = $normalizedItemName.Substring($normalizedSetName.Length).Trim()
                if ($remainingText.Length -eq 0 -or $remainingText -match '^[\s\W]') {
                    return "set"
                }
            }
        }
    }
    
    # Si aucune qualité n'a été détectée, retour par défaut WHITE
    return "white"
}

# ================================================================
# DISCORD WEBHOOKS
# ================================================================

function Get-ItemColor {
    param([string]$Quality, [hashtable]$Config)
    
    $colorValue = 12370112
    
    switch -Regex ($Quality) {
        "^unique$" { $colorValue = if ($Config.couleurs.unique) { $Config.couleurs.unique } else { 16753920 } }
        "^set$" { $colorValue = if ($Config.couleurs.set) { $Config.couleurs.set } else { 65280 } }
        "^rare$" { $colorValue = if ($Config.couleurs.rare) { $Config.couleurs.rare } else { 16776960 } }
        "^magic$" { $colorValue = if ($Config.couleurs.magic) { $Config.couleurs.magic } else { 1146986 } }
        "^rune$" { $colorValue = if ($Config.couleurs.rune) { $Config.couleurs.rune } else { 16753920 } }
        "^high_rune$" { $colorValue = if ($Config.couleurs.high_rune) { $Config.couleurs.high_rune } else { 16711680 } }
        "^uber_keys$" { $colorValue = if ($Config.couleurs.uber_keys) { $Config.couleurs.uber_keys } else { 8388863 } }
        "^essences$" { $colorValue = if ($Config.couleurs.essences) { $Config.couleurs.essences } else { 65535 } }
        "^misc$" { $colorValue = if ($Config.couleurs.misc) { $Config.couleurs.misc } else { 8421504 } }
        "^white$" { $colorValue = 12370112 }
        "^runeword$" { $colorValue = if ($Config.couleurs.runeword) { $Config.couleurs.runeword } else { 16750848 } }
        "^crafted$" { $colorValue = if ($Config.couleurs.crafted) { $Config.couleurs.crafted } else { 16744448 } }
        "_unique$" { $colorValue = if ($Config.couleurs.unique) { $Config.couleurs.unique } else { 16753920 } }
        "_magic$" { $colorValue = if ($Config.couleurs.magic) { $Config.couleurs.magic } else { 1146986 } }
        "_rare$" { $colorValue = if ($Config.couleurs.rare) { $Config.couleurs.rare } else { 16776960 } }
        "_white$" { $colorValue = 12370112 }
    }
    
    return $colorValue
}

function Get-WebhookForQuality {
    param([string]$Quality, [hashtable]$Config)
    
    function Is-ValidWebhook {
        param([string]$Url)
        if (-not $Url -or $Url -eq "" -or $Url -eq "false" -or $Url -eq "FALSE") {
            return $false
        }

        return $true
    }
    
    if (-not $Config.webhooks) {
        return $Config.discord_webhook
    }
    
    if ($Quality -match '^(small_charm|large_charm|grand_charm|jewel)_(.+)$') {
        $itemType = $Matches[1]
        $subQuality = $Matches[2]
        
        if ($Config.webhooks.$itemType -and (Is-ValidWebhook $Config.webhooks.$itemType)) {
            return $Config.webhooks.$itemType
        }
        
        $subWebhook = switch ($subQuality) {
            "unique" { $Config.webhooks.unique }
            "rare" { $Config.webhooks.rare }
            "magic" { $Config.webhooks.magic }
            "white" { $Config.webhooks.white }
        }
        
        if ($subWebhook -and (Is-ValidWebhook $subWebhook)) {
            return $subWebhook
        }
        
        return $Config.discord_webhook
    }
    
    $webhookUrl = switch ($Quality) {
        "unique" { $Config.webhooks.unique }
        "set" { $Config.webhooks.set }
        "rare" { $Config.webhooks.rare }
        "magic" { $Config.webhooks.magic }
        "rune" { $Config.webhooks.rune }
        "high_rune" { $Config.webhooks.high_rune }
        "uber_keys" { $Config.webhooks.uber_keys }
        "essences" { $Config.webhooks.essences }
        "misc" { $Config.webhooks.misc }
        "white" { $Config.webhooks.white }
        "runeword" { $Config.webhooks.runeword }
        "crafted" { $Config.webhooks.crafted }
    }
    
    if ($webhookUrl -and (Is-ValidWebhook $webhookUrl)) {
        return $webhookUrl
    }
    
    return $Config.discord_webhook
}

function Get-QualityEmoji {
    param([string]$Quality)
    
    switch -Regex ($Quality) {
        "^unique$" { return "🟠 **UNIQUE**" }
        "^set$" { return "🟢 **SET**" }
        "^rare$" { return "🟡 **RARE**" }
        "^magic$" { return "🔵 **MAGIC**" }
        "^rune$" { return "🟥 **RUNE**" }
        "^high_rune$" { return "🔴 **HIGH RUNE**" }
        "^uber_keys$" { return "🔑 **UBER KEY**" }
        "^essences$" { return "💎 **ESSENCE**" }
        "^white$" { return "⚪ **WHITE**" }
        "^runeword$" { return "🟧 **RUNEWORD**" }
        "^crafted$" { return "🟨 **CRAFTED**" }
        "small_charm_unique$" { return "🟠 **UNIQUE SMALL CHARM**" }
        "small_charm_magic$" { return "🔵 **MAGIC SMALL CHARM**" }
        "large_charm_unique$" { return "🟠 **UNIQUE LARGE CHARM**" }
        "large_charm_magic$" { return "🔵 **MAGIC LARGE CHARM**" }
        "grand_charm_unique$" { return "🟠 **UNIQUE GRAND CHARM**" }
        "grand_charm_magic$" { return "🔵 **MAGIC GRAND CHARM**" }
        "jewel_unique$" { return "🟠 **UNIQUE JEWEL**" }
        "jewel_rare$" { return "🟡 **RARE JEWEL**" }
        "jewel_magic$" { return "🔵 **MAGIC JEWEL**" }
        default { return "⚫ **UNKNOWN**" }
    }
}


function Get-QualityEmoji {
    param([string]$Quality)
    
    $emojiMap = @{
        "unique" = "🟠"
        "set" = "🟢"
        "rare" = "🟡"
        "magic" = "🔵"
        "rune" = "🟥"
        "high_rune" = "🔴"
        "uber_keys" = "🔑"
        "essences" = "💎"
        "misc" = "📦"
        "crafted" = "🟨"
        "small_charm_magic" = "🔵"
        "large_charm_magic" = "🔵"
        "grand_charm_magic" = "🔵"
        "small_charm_unique" = "🟠"
        "large_charm_unique" = "🟠"
        "grand_charm_unique" = "🟠"
        "small_charm_rare" = "🟡"
        "large_charm_rare" = "🟡"
        "grand_charm_rare" = "🟡"
        "jewel_magic" = "🔵"
        "jewel_rare" = "🟡"
        "white" = "⚪"
    }
    
    if ($emojiMap.ContainsKey($Quality)) {
        return $emojiMap[$Quality]
    }
    
    return "❔"
}

function Get-QualityLabel {
    param([string]$Quality)
    
    $labelMap = @{
        "unique" = "UNIQUE"
        "set" = "SET"
        "rare" = "RARE"
        "magic" = "MAGIC"
        "rune" = "RUNE"
        "high_rune" = "HIGH RUNE"
        "crafted" = "CRAFTED"
        "small_charm_magic" = "MAGIC SMALL CHARM"
        "large_charm_magic" = "MAGIC LARGE CHARM"
        "grand_charm_magic" = "MAGIC GRAND CHARM"
        "small_charm_unique" = "UNIQUE SMALL CHARM"
        "large_charm_unique" = "UNIQUE LARGE CHARM"
        "grand_charm_unique" = "UNIQUE GRAND CHARM"
        "small_charm_rare" = "RARE SMALL CHARM"
        "large_charm_rare" = "RARE LARGE CHARM"
        "grand_charm_rare" = "RARE GRAND CHARM"
        "jewel_magic" = "MAGIC JEWEL"
        "jewel_rare" = "RARE JEWEL"
        "misc" = "MISC"
        "white" = "NORMAL"
    }
    
    if ($labelMap.ContainsKey($Quality)) {
        return $labelMap[$Quality]
    }
    
    return "UNKNOWN"
}

function Send-ToDiscord {
    param(
        [string]$WebhookUrl,
        [hashtable]$Item,
        [string]$InvFile,
        [int]$Color,
        [string]$Location,
        [string]$Quality,
        [hashtable]$Config
    )

    if (-not $WebhookUrl -or $WebhookUrl -eq "" -or $WebhookUrl -eq "false" -or $WebhookUrl -eq "FALSE") {
        Write-Host "  ⏭ No webhook configured for quality: $Quality" -ForegroundColor Yellow
        return
    }

    $imageUrl = Get-ImageUrl -InvFile $InvFile

    try {
        $dt = Parse-Timestamp -Timestamp $Item.Timestamp
        $timestamp12hr = $dt.ToString("hh:mm:ss tt")
    }
    catch {
        $timestamp12hr = $Item.Timestamp
    }

    # Vérifier si on affiche la qualité (par défaut: true)
    $showQuality = if ($Config -and $Config.ContainsKey("show_quality")) { $Config.show_quality } else { $true }

    # Format demandé : Quality → Localisation → Required Level → Ligne vide → Stats → Ligne vide → Found by
    $qualityText = ""
    if ($showQuality) {
        $qualityEmoji = Get-QualityEmoji -Quality $Quality
        $qualityLabel = Get-QualityLabel -Quality $Quality
        $qualityText = "$qualityEmoji **$qualityLabel**`n`n"
    }
    $locationText = if ($Location) { "📍 **$Location**`n`n" } else { "" }
    
    # Combiner Type + Stats pour avoir toutes les infos
    $allStats = if ($Item.Type) { "$($Item.Type)`n$($Item.Stats)" } else { $Item.Stats }
    
    # Extraire Required Level de toutes les stats
    $statsLines = $allStats -split "`n"
    $requiredLevel = $null
    $otherStats = @()
    
    foreach ($line in $statsLines) {
        if ($line -match '^Required Level:' -or $line -match '^Level Required:' -or $line -match 'pc requiert') {
            $requiredLevel = $line.Trim()
        }
        else {
            $otherStats += $line
        }
    }
    
    # Construire le message avec lignes vides
    $requiredText = if ($requiredLevel) { "$requiredLevel`n`n" } else { "" }
    $remainingStats = ($otherStats | Where-Object { $_ }) -join "`n"
    
    $description = "$qualityText$locationText$requiredText$remainingStats`n`n*Found by: $($Item.FoundBy) | $timestamp12hr*"
    $embed = @{
        color = $Color
        title = $Item.Name
        description = $description
        thumbnail = @{ url = $imageUrl }
    }
    
    $payload = @{ embeds = @($embed) } | ConvertTo-Json -Depth 10
    
    try {
        Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $payload -ContentType "application/json; charset=utf-8" -TimeoutSec 10
        $locDisplay = if ($Location) { " in $Location" } else { "" }
        Write-Host "  ✓ Sent: $($Item.Name)$locDisplay" -ForegroundColor Green
    }
    catch {
        Write-Host "  ❌ Discord error: $_" -ForegroundColor Red
    }
}

# ================================================================
# LOCATION TRACKING
# ================================================================

function Get-LastScene {
    param(
        [string]$LogsPath,
        [string]$Timestamp,
        [string]$PlayerName
    )
    
    try {
        $itemDateTime = Parse-Timestamp -Timestamp $Timestamp

        # Get the TagId for this character
        $tagId = $global:nameToTagId[$PlayerName.ToLower()]

        # DEBUG: Afficher le mapping
        Write-Host "    [DEBUG] PlayerName: $PlayerName, TagId: $tagId, ItemTime: $($itemDateTime.ToString('HH:mm:ss'))" -ForegroundColor DarkGray

        # Get recent log files
        $logFiles = Get-ChildItem -Path $LogsPath -Filter "*.txt" -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 15

        if (-not $logFiles) {
            # Write-Host "    [DEBUG] No log files found in $LogsPath" -ForegroundColor DarkGray
            return $null
        }

        # Try to find the log file for this specific TagId
        $playerLogFile = $null
        if ($tagId) {
            foreach ($logFile in $logFiles) {
                $firstLines = Get-Content $logFile.FullName -TotalCount 20 -ErrorAction SilentlyContinue
                foreach ($line in $firstLines) {
                    if ($line -match "CharacterName:(\w+)") {
                        $logTagId = $Matches[1]
                        if ($logTagId -eq $tagId) {
                            $playerLogFile = $logFile
                            break
                        }
                    }
                }
                if ($playerLogFile) {
                    break
                }
            }
        }

        # If we didn't find a specific log, use the most recent logs as fallback
        $logsToSearch = if ($playerLogFile) { @($playerLogFile) } else { $logFiles | Select-Object -First 3 }

        # DEBUG: Afficher si on a trouvé le bon fichier
        if ($playerLogFile) { Write-Host "    [DEBUG] Log found: $($playerLogFile.Name)" -ForegroundColor DarkGray }
        else { Write-Host "    [DEBUG] No specific log found for $PlayerName (tagId=$tagId), using fallback" -ForegroundColor Yellow }
        
        $lastScene = $null
        $lastNonTownScene = $null
        $isGambled = $false
        $isCrafted = $false
        
        foreach ($logFile in $logsToSearch) {
            $content = Get-Content $logFile.FullName -ErrorAction SilentlyContinue
            
            foreach ($line in $content) {
                # Check for gambling messages
                # GID logs are in UTC, so we convert them to local time with -IsUTC
                if ($line -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}).*Item .* has been gambled') {
                    $logTimestamp = $Matches[1]
                    try {
                        $logDateTime = Parse-Timestamp -Timestamp $logTimestamp -IsUTC
                        # Check if gambling happened within 10 seconds before the item timestamp
                        $timeDiff = ($itemDateTime - $logDateTime).TotalSeconds
                        if ($timeDiff -ge 0 -and $timeDiff -le 10) {
                            $isGambled = $true
                        }
                    }
                    catch { }
                }

                # Check for crafting messages
                # GID logs are in UTC, so we convert them to local time with -IsUTC
                if ($line -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}).*The result of crafting is') {
                    $logTimestamp = $Matches[1]
                    try {
                        $logDateTime = Parse-Timestamp -Timestamp $logTimestamp -IsUTC
                        # Check if crafting happened within 10 seconds before the item timestamp
                        $timeDiff = ($itemDateTime - $logDateTime).TotalSeconds
                        if ($timeDiff -ge 0 -and $timeDiff -le 10) {
                            $isCrafted = $true
                        }
                    }
                    catch { }
                }

                # Check for scene changes
                # GID logs are in UTC, so we convert them to local time with -IsUTC
                if ($line -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}).*In the scene (\d+)') {
                    $logTimestamp = $Matches[1]
                    $sceneId = $Matches[2]

                    try {
                        $logDateTime = Parse-Timestamp -Timestamp $logTimestamp -IsUTC

                        if ($logDateTime -le $itemDateTime) {
                            $lastScene = $sceneId

                            # Keep track of last non-town scene
                            if ($global:townScenes -notcontains $sceneId) {
                                $lastNonTownScene = $sceneId
                            }
                        }
                    }
                    catch { }
                }
            }
        }
        
        # If item was gambled or crafted, return that instead
        if ($isGambled) {
            return "GAMBLED"
        }
        
        if ($isCrafted) {
            return "CRAFTED"
        }
        
        # If we're in a town, return the last non-town scene (where they likely found the item)
        if ($lastScene -and $global:townScenes -contains $lastScene) {
            if ($lastNonTownScene) {
                return $lastNonTownScene
            }
        }

        # DEBUG: Si pas de scene trouvée
        if (-not $lastScene) {
            Write-Host "    ⚠ [LOCATION] No scene found - Item: $($itemDateTime.ToString('HH:mm:ss')), LogFile: $(if($playerLogFile){$playerLogFile.Name}else{'fallback'})" -ForegroundColor DarkYellow
        }

        return $lastScene
    }
    catch {
        Write-Host "    ⚠ [LOCATION] Error: $_" -ForegroundColor DarkYellow
        return $null
    }
}

# ================================================================
# DEATH TRACKING & AUTO-KILL
# ================================================================

function Get-PIDFromLog {
    param([string]$LogPath)
    
    if (-not (Test-Path $LogPath)) { return $null }
    
    try {
        $firstLines = Get-Content $LogPath -TotalCount 30 -ErrorAction SilentlyContinue
        $processPid = $null
        $characterName = $null
        
        foreach ($line in $firstLines) {
            # Chercher "Start to Diablo 2 (PID)"
            if ($line -match 'Start to Diablo 2 \((\d+)\)') { 
                $processPid = [int]$Matches[1] 
            }
            # Chercher le TagId
            if ($line -match 'CharacterName:(\w+)') { 
                $characterName = $Matches[1] 
            }
            if ($processPid -and $characterName) { 
                break 
            }
        }
        
        if ($processPid -and $characterName) {
            return @{ 
                PID = $processPid
                CharacterName = $characterName
                LogPath = $LogPath 
            }
        }
    }
    catch { }
    
    return $null
}

function Update-PIDMapping {
    param([hashtable]$Config)

    if (-not $Config.logs_path -or -not (Test-Path $Config.logs_path)) { return }

    try {
        # 1. Nettoyer les anciens PIDs qui ne sont plus valides (crash/kill/restart)
        $keysToRemove = @()
        foreach ($charName in @($global:pidMapping.Keys)) {
            $oldPid = $global:pidMapping[$charName].PID
            $processStillRunning = Get-Process -Id $oldPid -ErrorAction SilentlyContinue
            if (-not $processStillRunning) {
                $keysToRemove += $charName
            }
        }
        foreach ($key in $keysToRemove) {
            $deadPid = $global:pidMapping[$key].PID
            Write-Host "  ⚫ [PID] $key - PID $deadPid removed (dead)" -ForegroundColor Yellow
            $global:pidMapping.Remove($key)
        }

        # 2. Scanner les logs récents pour trouver les nouveaux PIDs
        $logFiles = Get-ChildItem -Path $Config.logs_path -Filter "*.txt" -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 20

        foreach ($logFile in $logFiles) {
            $pidInfo = Get-PIDFromLog -LogPath $logFile.FullName

            if ($pidInfo) {
                $charName = $pidInfo.CharacterName
                $newPid = $pidInfo.PID

                $process = Get-Process -Id $newPid -ErrorAction SilentlyContinue

                if ($process) {
                    # Vérifier si c'est un nouveau PID ou un PID différent (restart)
                    $isNew = -not $global:pidMapping.ContainsKey($charName)
                    $isChanged = $false
                    $timestamp = Get-Date -Format "HH:mm:ss"

                    if (-not $isNew) {
                        $oldPid = $global:pidMapping[$charName].PID
                        if ($oldPid -ne $newPid) {
                            $isChanged = $true
                            Write-Host "  🔄 [PID] $charName - Restarted: $oldPid → $newPid" -ForegroundColor Cyan
                        }
                    }

                    if ($isNew) {
                        Write-Host "  🟢 [PID] $charName - New session: PID $newPid" -ForegroundColor Green
                    }

                    $global:pidMapping[$charName] = @{
                        PID = $newPid
                        LogPath = $logFile.FullName
                        LastUpdate = Get-Date
                    }
                }
            }
        }
    }
    catch { }
}

function Close-D2RInstance {
    param(
        [int]$ProcessPID,
        [string]$CharacterName,
        [int]$DeathCount,
        [string]$Reason = "Unknown"
    )

    Write-Host "`n============================================" -ForegroundColor Red
    Write-Host "  🚨 NUCLEAR KILL MODE ACTIVATED 🚨" -ForegroundColor Yellow
    Write-Host "============================================" -ForegroundColor Red
    Write-Host "  Target: $CharacterName" -ForegroundColor White
    Write-Host "  PID: $ProcessPID" -ForegroundColor White
    Write-Host "  Deaths: $DeathCount" -ForegroundColor White
    Write-Host "  Reason: $Reason" -ForegroundColor White
    Write-Host "============================================`n" -ForegroundColor Red

    # Helper function to reset counters after successful kill
    function Reset-AfterKill {
        if ($global:pidMapping.ContainsKey($CharacterName)) {
            $global:pidMapping.Remove($CharacterName)
        }
        if ($global:consecutiveDeaths.ContainsKey($CharacterName)) {
            $global:consecutiveDeaths[$CharacterName] = 0
            Write-Host "  ✅ [COUNTER RESET] Counter reset to 0 for $CharacterName" -ForegroundColor Green
        }
    }

    # Check if process exists
    $process = Get-Process -Id $ProcessPID -ErrorAction SilentlyContinue
    if (-not $process) {
        Write-Host "  ⚠ Process not found - may already be dead" -ForegroundColor Yellow
        if ($global:consecutiveDeaths.ContainsKey($CharacterName)) {
            $global:consecutiveDeaths[$CharacterName] = 0
            Write-Host "  ✅ [COUNTER RESET] Counter reset to 0 for $CharacterName (process already dead)" -ForegroundColor Green
        }
        return $true
    }

    # METHOD 1: Graceful close
    Write-Host "  ⚡ METHOD 1: Attempting graceful close..." -ForegroundColor Cyan
    try {
        $process.CloseMainWindow() | Out-Null
        Start-Sleep -Seconds 2
        $process = Get-Process -Id $ProcessPID -ErrorAction SilentlyContinue
        if (-not $process) {
            Write-Host "  ✓ Method 1 succeeded - Process closed gracefully" -ForegroundColor Green
            Reset-AfterKill
            return $true
        }
        Write-Host "  ⚠ Method 1 failed - Process still running" -ForegroundColor Yellow
    }
    catch {
        Write-Host "  ⚠ Method 1 error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # METHOD 2: Stop-Process
    Write-Host "  ⚡ METHOD 2: Stop-Process (standard kill)..." -ForegroundColor Cyan
    try {
        Stop-Process -Id $ProcessPID -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        $process = Get-Process -Id $ProcessPID -ErrorAction SilentlyContinue
        if (-not $process) {
            Write-Host "  ✓ Method 2 succeeded - Process force killed" -ForegroundColor Green
            Reset-AfterKill
            return $true
        }
        Write-Host "  ⚠ Method 2 failed - Process still running" -ForegroundColor Yellow
    }
    catch {
        Write-Host "  ⚠ Method 2 error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # METHOD 3: WMI Terminate
    Write-Host "  ⚡ METHOD 3: WMI Terminate..." -ForegroundColor Cyan
    try {
        $wmiProcess = Get-WmiObject -Class Win32_Process -Filter "ProcessId = $ProcessPID" -ErrorAction SilentlyContinue
        if ($wmiProcess) {
            $wmiProcess.Terminate() | Out-Null
            Start-Sleep -Seconds 2
        }
        $process = Get-Process -Id $ProcessPID -ErrorAction SilentlyContinue
        if (-not $process) {
            Write-Host "  ✓ Method 3 succeeded - WMI terminate" -ForegroundColor Green
            Reset-AfterKill
            return $true
        }
        Write-Host "  ⚠ Method 3 failed - Process still running" -ForegroundColor Yellow
    }
    catch {
        Write-Host "  ⚠ Method 3 error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # METHOD 4: taskkill /F (no elevation)
    Write-Host "  ⚡ METHOD 4: taskkill /F..." -ForegroundColor Cyan
    try {
        $result = Start-Process "taskkill" -ArgumentList "/F /PID $ProcessPID" -NoNewWindow -Wait -PassThru -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        $process = Get-Process -Id $ProcessPID -ErrorAction SilentlyContinue
        if (-not $process) {
            Write-Host "  ✓ Method 4 succeeded - taskkill force" -ForegroundColor Green
            Reset-AfterKill
            return $true
        }
        Write-Host "  ⚠ Method 4 failed - Process still running" -ForegroundColor Yellow
    }
    catch {
        Write-Host "  ⚠ Method 4 error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # METHOD 5: taskkill with elevation (Run As Admin)
    Write-Host "  ⚡ METHOD 5: taskkill /F with elevation (Admin)..." -ForegroundColor Cyan
    try {
        Start-Process "taskkill" -ArgumentList "/F /PID $ProcessPID" -Verb RunAs -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        $process = Get-Process -Id $ProcessPID -ErrorAction SilentlyContinue
        if (-not $process) {
            Write-Host "  ✓ Method 5 succeeded - taskkill with elevation" -ForegroundColor Green
            Reset-AfterKill
            return $true
        }
        Write-Host "  ⚠ Method 5 failed - Process still running" -ForegroundColor Yellow
    }
    catch {
        Write-Host "  ⚠ Method 5 error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # METHOD 6: NtTerminateProcess (kernel-mode)
    Write-Host "  ⚡ METHOD 6: NtTerminateProcess (kernel-mode)..." -ForegroundColor Cyan
    try {
        $handle = [System.Diagnostics.Process]::GetProcessById($ProcessPID).Handle
        $ntdll = Add-Type -MemberDefinition @"
        [DllImport("ntdll.dll", SetLastError=true)]
        public static extern int NtTerminateProcess(IntPtr ProcessHandle, uint ExitStatus);
"@ -Name "NtDll$ProcessPID" -Namespace "Win32" -PassThru -ErrorAction SilentlyContinue
        if ($ntdll) {
            $ntdll::NtTerminateProcess($handle, 1) | Out-Null
            Start-Sleep -Seconds 2
        }
        $process = Get-Process -Id $ProcessPID -ErrorAction SilentlyContinue
        if (-not $process) {
            Write-Host "  ✓ Method 6 succeeded - Nuclear-level termination" -ForegroundColor Green
            Reset-AfterKill
            return $true
        }
        Write-Host "  ⚠ Method 6 failed - Process still running" -ForegroundColor Yellow
    }
    catch {
        Write-Host "  ⚠ Method 6 error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # METHOD 7: WMIC process delete (alternative)
    Write-Host "  ⚡ METHOD 7: WMIC process delete..." -ForegroundColor Cyan
    try {
        $wmicResult = & wmic process where "ProcessId=$ProcessPID" delete 2>&1
        Start-Sleep -Seconds 2
        $process = Get-Process -Id $ProcessPID -ErrorAction SilentlyContinue
        if (-not $process) {
            Write-Host "  ✓ Method 7 succeeded - WMIC delete" -ForegroundColor Green
            Reset-AfterKill
            return $true
        }
        Write-Host "  ⚠ Method 7 failed - Process still running" -ForegroundColor Yellow
    }
    catch {
        Write-Host "  ⚠ Method 7 error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    Write-Host "  ✗ ALL METHODS FAILED - Process is unkillable!" -ForegroundColor Red
    Write-Host "  ⚠ You need to run PowerShell as Administrator to kill D2R processes" -ForegroundColor Yellow
    Write-Host "  ⚠ Or manually kill the process in Task Manager" -ForegroundColor Yellow

    # Ne PAS reset le compteur si on n'a pas réussi à tuer le processus
    # Ainsi le script réessaiera au prochain cycle
    return $false
}


function Send-DeathAlert {
    param(
        [string]$WebhookUrl,
        [string]$CharacterName,
        [string]$TagId,
        [int]$DeathCount,
        [string]$Location,
        [bool]$IncludeTagId
    )
    
    if (-not $WebhookUrl -or $WebhookUrl -eq "" -or $WebhookUrl -eq "false" -or $WebhookUrl -eq "FALSE" -or $WebhookUrl -eq "WEBHOOK OUT TAG ID") {
        return
    }
    
    try {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $locationText = if ($Location) { "in **$Location**" } else { "" }
        $tagIdText = if ($IncludeTagId -and $TagId) { "`n🏷️ TagId: **$TagId**" } else { "" }
        
        $description = "💀 **$CharacterName** is dead! $locationText$tagIdText`n`n🔢 Deaths today: **$DeathCount**`n⏰ Time: $timestamp"
        
        $embed = @{
            color = 16711680
            title = "⚠️ Character Death Alert"
            description = $description
        }
        
        $payload = @{ embeds = @($embed) } | ConvertTo-Json -Depth 10
        
        Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $payload -ContentType "application/json; charset=utf-8" -TimeoutSec 10
        
        $tagDisplay = if ($IncludeTagId) { "with TagId" } else { "without TagId" }
        Write-Host "  💀 Death alert sent for $CharacterName [$tagDisplay] (Total: $DeathCount)" -ForegroundColor Red
    }
    catch {
        Write-Host "  Error sending death alert: $_" -ForegroundColor Red
    }
}

# ================================================================
# INITIALISATION : MARQUER LES MORTS EXISTANTES COMME DÉJÀ VUES
# ================================================================
function Initialize-ReportedDeaths {
    param([hashtable]$Config)

    if (-not $Config.track_deaths -or -not $Config.logs_path) { return }

    Write-Host "  📋 Initializing death cache from existing logs..." -ForegroundColor Cyan
    $count = 0

    try {
        $logFiles = Get-ChildItem -Path $Config.logs_path -Filter "*.txt" -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 30

        if (-not $logFiles) { return }

        foreach ($logFile in $logFiles) {
            try {
                $content = Get-Content $logFile.FullName -Tail 100 -ErrorAction SilentlyContinue
                if (-not $content) { continue }

                $tagId = $null
                $characterName = $null
                $firstLines = Get-Content $logFile.FullName -TotalCount 20 -ErrorAction SilentlyContinue
                foreach ($line in $firstLines) {
                    if ($line -match "CharacterName:(\w+)") {
                        $tagId = $Matches[1]
                        if ($global:charMapping.ContainsKey($tagId)) {
                            $characterName = $global:charMapping[$tagId]
                        }
                        break
                    }
                }
                $uniqueId = if ($tagId) { $tagId } else { $characterName }
                if (-not $characterName) {
                    if ($tagId) { $characterName = $tagId }
                    else {
                        $characterName = [System.IO.Path]::GetFileNameWithoutExtension($logFile.Name)
                        $uniqueId = $characterName
                    }
                }

                foreach ($line in $content) {
                    if ($line -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),\d+.*Character is dead') {
                        $deathKey = "$uniqueId-$($Matches[1])"
                        if (-not $global:reportedDeaths.ContainsKey($deathKey)) {
                            $global:reportedDeaths[$deathKey] = $true
                            $count++
                        }
                    }
                }
            }
            catch { }
        }
        Write-Host "  ✓ Death cache ready: $count existing deaths marked as seen (no notification sent)" -ForegroundColor Green
    }
    catch { }
}

# ================================================================
# [CORRECTION] FONCTION DE RESET AUTOMATIQUE DES DEATHS TODAY
# ================================================================
function Check-And-Reset-DeathsToday {
    param([hashtable]$Config)
    
    $currentTime = Get-Date
    $inactivityThreshold = [TimeSpan]::FromHours(11)
    
    # Parcourir tous les compteurs de deaths
    $keysToCheck = @($global:lastDeathTimestamps.Keys)
    
    foreach ($uniqueId in $keysToCheck) {
        if ($global:lastDeathTimestamps.ContainsKey($uniqueId)) {
            $lastDeathTime = $global:lastDeathTimestamps[$uniqueId]
            $timeSinceLastDeath = $currentTime - $lastDeathTime
            
            # Si plus de 11 heures depuis la dernière mort, reset le compteur
            if ($timeSinceLastDeath -ge $inactivityThreshold) {
                if ($global:deathCount.ContainsKey($uniqueId)) {
                    $oldCount = $global:deathCount[$uniqueId]
                    $global:deathCount[$uniqueId] = 0
                    Write-Host "  🔄 [AUTO-RESET] Deaths today for $uniqueId reset after 11h inactivity ($oldCount → 0)" -ForegroundColor Cyan
                }
                
                # Supprimer le timestamp pour éviter de le vérifier à nouveau
                $global:lastDeathTimestamps.Remove($uniqueId)
            }
        }
    }
}

function Check-ForDeaths {
    param([hashtable]$Config)
    
    if (-not $Config.track_deaths -or (-not $Config.death_webhook_with_id -and -not $Config.death_webhook_without_id) -or -not $Config.logs_path) {
        return
    }
    
    # [CORRECTION] Vérifier et reset les compteurs "Deaths today" si inactifs depuis 11h
    Check-And-Reset-DeathsToday -Config $Config
    
    Update-PIDMapping -Config $Config
    
    try {
        $logFiles = Get-ChildItem -Path $Config.logs_path -Filter "*.txt" -ErrorAction SilentlyContinue | 
                    Sort-Object LastWriteTime -Descending | 
                    Select-Object -First 10
        
        if (-not $logFiles) { return }
        
        foreach ($logFile in $logFiles) {
            try {
                $content = Get-Content $logFile.FullName -Tail 100 -ErrorAction SilentlyContinue
                if (-not $content) { continue }
                
                $characterName = $null
                $tagId = $null
                
                $firstLines = Get-Content $logFile.FullName -TotalCount 20 -ErrorAction SilentlyContinue
                foreach ($line in $firstLines) {
                    if ($line -match "CharacterName:(\w+)") {
                        $tagId = $Matches[1]
                        # Le mapping est: $global:charMapping[$tagId] = $charName
                        # Donc on cherche directement avec le TagId comme clé
                        if ($global:charMapping.ContainsKey($tagId)) {
                            $characterName = $global:charMapping[$tagId]
                        }
                        break
                    }
                }
                
                $uniqueId = if ($tagId) { $tagId } else { $characterName }
                
                if (-not $characterName) {
                    if ($tagId) {
                        $characterName = $tagId
                    } else {
                        $characterName = [System.IO.Path]::GetFileNameWithoutExtension($logFile.Name)
                        $uniqueId = $characterName
                    }
                }
                
                foreach ($line in $content) {
                    if ($line -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),\d+.*Character is dead') {
                        $deathTimestamp = $Matches[1]
                        $deathKey = "$uniqueId-$deathTimestamp"

                        if (-not $global:reportedDeaths.ContainsKey($deathKey)) {
                            if (-not $global:deathCount.ContainsKey($uniqueId)) {
                                $global:deathCount[$uniqueId] = 0
                            }
                            $global:deathCount[$uniqueId]++

                            # [CORRECTION] Enregistrer le timestamp de cette mort pour le reset automatique
                            # GID logs are in UTC, so we convert them to local time with -IsUTC
                            try {
                                $deathDateTime = Parse-Timestamp -Timestamp $deathTimestamp -IsUTC
                                if ($deathDateTime) {
                                    $global:lastDeathTimestamps[$uniqueId] = $deathDateTime
                                }
                            }
                            catch { }
                            
                            $location = $null
                            $lastScene = $null
                            
                            for ($i = $content.Count - 1; $i -ge 0; $i--) {
                                $logLine = $content[$i]

                                if ($logLine -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),\d+.*In the scene (\d+)') {
                                    $sceneTimestamp = $Matches[1]
                                    $sceneId = $Matches[2]

                                    try {
                                        # GID logs are in UTC, so we convert them to local time with -IsUTC
                                        $sceneDateTime = Parse-Timestamp -Timestamp $sceneTimestamp -IsUTC
                                        $deathDateTime = Parse-Timestamp -Timestamp $deathTimestamp -IsUTC

                                        if ($sceneDateTime -le $deathDateTime) {
                                            $lastScene = $sceneId
                                            break
                                        }
                                    }
                                    catch { }
                                }
                            }
                            
                            if ($lastScene) {
                                $location = Get-SceneName -SceneId $lastScene
                            }
                            
                            $displayTime = Get-Date -Format "HH:mm:ss"
                            Write-Host "[$displayTime] 💀 Death detected: $characterName [$tagId]" -ForegroundColor Red
                            
                            if ($Config.death_webhook_with_id) {
                                Send-DeathAlert -WebhookUrl $Config.death_webhook_with_id `
                                              -CharacterName $characterName `
                                              -TagId $tagId `
                                              -DeathCount $global:deathCount[$uniqueId] `
                                              -Location $location `
                                              -IncludeTagId $true
                            }
                            
                            if ($Config.death_webhook_without_id) {
                                Send-DeathAlert -WebhookUrl $Config.death_webhook_without_id `
                                              -CharacterName $characterName `
                                              -TagId $tagId `
                                              -DeathCount $global:deathCount[$uniqueId] `
                                              -Location $location `
                                              -IncludeTagId $false
                            }
                            
                            $global:reportedDeaths[$deathKey] = $true

                            # AUTO-KILL LOGIC - Smart Counter Management
                            if (-not $global:consecutiveDeaths.ContainsKey($uniqueId)) {
                                $global:consecutiveDeaths[$uniqueId] = 0
                            }
                            
                            # Wait for log to be written (only if this is first death or counter is 0)
                            if ($global:consecutiveDeaths[$uniqueId] -eq 0) {
                                Write-Host "  ⏳ Waiting 6 seconds for game recreation log..." -ForegroundColor Cyan
                                Start-Sleep -Seconds 6
                            }
                            
                            # Reload log content after waiting to get latest lines (including "Create a new game")
                            if ($global:consecutiveDeaths[$uniqueId] -eq 0) {
                                $content = Get-Content $logFile.FullName -Tail 100 -ErrorAction SilentlyContinue
                                Write-Host "  📄 Log reloaded after wait" -ForegroundColor Gray
                            }
                            
                            # Check if game was exited/recreated after death (normal death, not death bug)
                            $gameRecreated = $false
                            $deathLineIndex = -1
                            
                            # Find the death line index in the log
                            for ($i = 0; $i -lt $content.Count; $i++) {
                                if ($content[$i] -match "Character is dead" -and $content[$i] -match $deathTimestamp) {
                                    $deathLineIndex = $i
                                    break
                                }
                            }
                            
                            # Check lines AFTER death for game exit/recreation (within next 20 lines)
                            if ($deathLineIndex -ge 0) {
                                for ($i = $deathLineIndex + 1; $i -lt [Math]::Min($deathLineIndex + 20, $content.Count); $i++) {
                                    if ($content[$i] -match 'Create a new game') {
                                        $gameRecreated = $true
                                        if ($global:consecutiveDeaths[$uniqueId] -gt 0) {
                                            Write-Host "  ✅ [COUNTER RESET] $characterName - 'Create a new game' detected after death, consecutive deaths reset: $($global:consecutiveDeaths[$uniqueId]) → 0" -ForegroundColor Green
                                        } else {
                                            Write-Host "  ✅ $characterName - Normal death, 'Create a new game' detected (counter stays at 0)" -ForegroundColor Green
                                        }
                                        $global:consecutiveDeaths[$uniqueId] = 0
                                        break
                                    }
                                }
                            }
                            
                            # Increment counter ONLY if no game recreation detected (death bug scenario)
                            if (-not $gameRecreated) {
                                $global:consecutiveDeaths[$uniqueId]++
                                $threshold = if ($Config.death_bug_threshold) { $Config.death_bug_threshold } else { 30 }
                                Write-Host "  ⚠️ [DEATH BUG ALERT] No game exit/recreation detected after death! Counter: $($global:consecutiveDeaths[$uniqueId])/$threshold" -ForegroundColor Yellow
                            }
                            
                            $deathThreshold = if ($Config.death_bug_threshold) { $Config.death_bug_threshold } else { 30 }
                            $consecutiveDeathsSnapshot = $global:consecutiveDeaths[$uniqueId]
                            if ($consecutiveDeathsSnapshot -ge $deathThreshold -and $Config.auto_kill_death_beug_enabled) {
                                Write-Host "`n⚠⚠⚠ DEATH BUG DETECTED ⚠⚠⚠" -ForegroundColor Red
                                Write-Host "  Character: $characterName | Deaths: $($global:consecutiveDeaths[$uniqueId])" -ForegroundColor Yellow
                                Write-Host "  Looking for UniqueId: $uniqueId" -ForegroundColor Yellow
                                
                                # [DEBUG] Afficher tous les mappings disponibles
                                if ($global:pidMapping.Count -gt 0) {
                                    Write-Host "  Available PID mappings:" -ForegroundColor Cyan
                                    foreach ($key in $global:pidMapping.Keys) {
                                        $mapping = $global:pidMapping[$key]
                                        Write-Host "    - $key → PID $($mapping.PID)" -ForegroundColor Cyan
                                    }
                                }
                                else {
                                    Write-Host "  ⚠ No PID mappings available at all!" -ForegroundColor Red
                                }
                                
                                if ($global:pidMapping.ContainsKey($uniqueId)) {
                                    $pidInfo = $global:pidMapping[$uniqueId]
                                    $processPid = $pidInfo.PID
                                    Write-Host "  Found PID: $processPid - Closing..." -ForegroundColor Yellow
                                    
                                    $closed = Close-D2RInstance -ProcessPID $processPid -CharacterName $uniqueId -DeathCount $global:consecutiveDeaths[$uniqueId] -Reason "Death bug"
                                    
                                    if ($closed -and $Config.auto_kill_death_beug) {
                                        $timestamp = Get-Date -Format "HH:mm:ss"
                                        $locationText = if ($location) { "in **$location**" } else { "" }
                                        $description = "💀 **$characterName** death bug detected! $locationText`n🏷️ TagId: **$tagId**`n`n💀 Consecutive Deaths: **$consecutiveDeathsSnapshot**`n🔢 Process ID: **$processPid**`n⚡ Action: **D2R instance terminated**`n⏰ Time: $timestamp"
                                        $embed = @{ color = 16711680; title = "🚨 AUTO-KILL: Death Bug"; description = $description }
                                        $payload = @{ content = "⚠️ **DEATH BUG AUTO-KILL** ⚠️"; embeds = @($embed) } | ConvertTo-Json -Depth 10
                                        try {
                                            Invoke-RestMethod -Uri $Config.auto_kill_death_beug -Method Post -Body $payload -ContentType "application/json; charset=utf-8" -TimeoutSec 10
                                            Write-Host "  ✓ Auto-kill alert sent" -ForegroundColor Green
                                        } catch { }
                                    }
                                } else {
                                    Write-Host "  ⚠ No PID mapping found for: $uniqueId" -ForegroundColor Yellow
                                    
                                    # [FALLBACK] Essayer de trouver un processus D2R manuellement
                                    Write-Host "  Attempting fallback: searching for D2R processes..." -ForegroundColor Cyan
                                    $d2rProcesses = Get-Process -Name "D2R" -ErrorAction SilentlyContinue
                                    
                                    if ($d2rProcesses) {
                                        Write-Host "  Found $($d2rProcesses.Count) D2R process(es):" -ForegroundColor Cyan
                                        foreach ($proc in $d2rProcesses) {
                                            Write-Host "    - PID: $($proc.Id), Started: $($proc.StartTime)" -ForegroundColor Cyan
                                        }
                                        
                                        # Si un seul processus D2R, on peut tenter de le killer
                                        if ($d2rProcesses.Count -eq 1) {
                                            $processPid = $d2rProcesses[0].Id
                                            Write-Host "  Only one D2R process found, attempting to kill PID: $processPid" -ForegroundColor Yellow
                                            
                                            $closed = Close-D2RInstance -ProcessPID $processPid -CharacterName $uniqueId -DeathCount $global:consecutiveDeaths[$uniqueId] -Reason "Death bug (fallback)"
                                            
                                            if ($closed -and $Config.auto_kill_death_beug) {
                                                $timestamp = Get-Date -Format "HH:mm:ss"
                                                $locationText = if ($location) { "in **$location**" } else { "" }
                                                $description = "💀 **$characterName** death bug detected! $locationText`n🏷️ TagId: **$tagId**`n`n💀 Consecutive Deaths: **$consecutiveDeathsSnapshot**`n🔢 Process ID: **$processPid** (fallback)`n⚡ Action: **D2R instance terminated**`n⏰ Time: $timestamp"
                                                $embed = @{ color = 16711680; title = "🚨 AUTO-KILL: Death Bug (Fallback)"; description = $description }
                                                $payload = @{ content = "⚠️ **DEATH BUG AUTO-KILL (FALLBACK)** ⚠️"; embeds = @($embed) } | ConvertTo-Json -Depth 10
                                                try {
                                                    Invoke-RestMethod -Uri $Config.auto_kill_death_beug -Method Post -Body $payload -ContentType "application/json; charset=utf-8" -TimeoutSec 10
                                                    Write-Host "  ✓ Auto-kill alert sent (fallback)" -ForegroundColor Green
                                                } catch { }
                                            }
                                        }
                                        else {
                                            Write-Host "  ⚠ Multiple D2R processes found, cannot determine which to kill" -ForegroundColor Yellow
                                            Write-Host "  Please check logs_path in config.json to ensure PID mapping works" -ForegroundColor Yellow
                                        }
                                    }
                                    else {
                                        Write-Host "  ⚠ No D2R processes found at all" -ForegroundColor Yellow
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch {
                Write-Host "  Error checking deaths in $($logFile.Name): $_" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "  Error in death check: $_" -ForegroundColor Red
    }
}

# ================================================================
# MONITORING
# ================================================================

function Start-Monitoring {
    param([hashtable]$Config)
    
    $lootPath = Join-Path $Config.root_path "Looted"
    
    
    $logFiles = @()
    if (Test-Path $lootPath) {
        $playerFolders = Get-ChildItem -Path $lootPath -Directory -ErrorAction SilentlyContinue
        
        foreach ($folder in $playerFolders) {
            $possibleNames = @("Looted", "Looted.txt", "Looted.log")
            foreach ($name in $possibleNames) {
                $logFile = Join-Path $folder.FullName $name
                if (Test-Path $logFile) {
                    $logFiles += Get-Item $logFile
                    break
                }
            }
        }
    }
    
    if ($logFiles.Count -eq 0) {
        Read-Host "Press Enter to exit"
        return
    }
    
    
    $tracking = @{}
    foreach ($log in $logFiles) {
        $playerName = Split-Path $log.Directory.Name -Leaf
        if (-not $playerName) { $playerName = $log.Directory.Name }
        
        $itemCount = Get-ItemCount -Path $log.FullName
        $tracking[$log.FullName] = @{
            Player = $playerName
            Count = $itemCount
            Size = (Get-Item $log.FullName).Length
        }
    }
    
    Write-Host "`n==========================================" -ForegroundColor Cyan
    Write-Host "  🟢 Monitoring active - Press Ctrl+C to stop" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Cyan
    
    $deathCheckCounter = 0
    $deathCheckInterval = 5
    
    while ($true) {
        Start-Sleep -Seconds $Config.sleep_seconds
        
        $deathCheckCounter++
        if ($deathCheckCounter -ge $deathCheckInterval) {
            Check-ForDeaths -Config $Config
            $deathCheckCounter = 0
        }
        
        foreach ($log in $logFiles) {
            if (-not (Test-Path $log.FullName)) { continue }
            
            $currentSize = (Get-Item $log.FullName).Length
            $track = $tracking[$log.FullName]
            
            if ($currentSize -ne $track.Size) {
                $newCount = Get-ItemCount -Path $log.FullName
                
                if ($newCount -gt $track.Count) {
                    $timestamp = Get-Date -Format "HH:mm:ss"
                    Write-Host "`n[$timestamp] 📦 New item for $($track.Player)" -ForegroundColor Cyan
                    
                    try {
                        $item = Get-ItemFromLog -Path $log.FullName -Index ($newCount - 1) -PlayerName $track.Player
                        
                        if ($item) {
                            $shouldSkip = $false
                            
                            if ($Config.objets_ignores) {
                                foreach ($skipItem in $Config.objets_ignores) {
                                    if ($item.Name -eq $skipItem) {
                                        Write-Host "  ⏭ Skipping: $($item.Name)" -ForegroundColor Yellow
                                        $shouldSkip = $true
                                        break
                                    }
                                }
                            }
                            
                            if (-not $shouldSkip) {
                                $quality = Get-ItemQuality -ItemName $item.Name -ItemType $item.Type -ItemStats $item.Stats -AllLines $item.AllLines
                                
                                $invFile = Get-GFX -ItemName $item.Name -ItemType $item.Type
                                $color = Get-ItemColor -Quality $quality -Config $Config
                                $webhookUrl = Get-WebhookForQuality -Quality $quality -Config $Config
                                
                                $location = $null
                                if ($Config.logs_path) {
                                    $sceneId = Get-LastScene -LogsPath $Config.logs_path -Timestamp $item.Timestamp -PlayerName $track.Player
                                    if ($sceneId) {
                                        $location = Get-SceneName -SceneId $sceneId
                                    }
                                }
                                
                                Write-Host "  📊 Quality: $quality - $($item.Name)$(if($location){" in $location"})" -ForegroundColor Gray
                                
                                Send-ToDiscord -WebhookUrl $webhookUrl -Item $item -InvFile $invFile -Color $color -Location $location -Quality $quality -Config $Config
                            }
                        }
                    }
                    catch {
                        Write-Host "  ❌ Error: $_" -ForegroundColor Red
                    }
                    
                    $track.Count = $newCount
                }
                $track.Size = $currentSize
            }
        }
    }
}

# ================================================================
# MAIN SCRIPT
# ================================================================

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     VERSION: FULL SUPPORT + CORRECTIONS" -ForegroundColor White
Write-Host "==========================================" -ForegroundColor Cyan

# Afficher les infos de timezone au démarrage
$localZone = [System.TimeZoneInfo]::Local
$utcNow = [DateTime]::UtcNow
$localNow = $utcNow.ToLocalTime()
Write-Host ""
Write-Host "  [TIMEZONE CONFIG]" -ForegroundColor Magenta
Write-Host "  PC Timezone: $($localZone.DisplayName)" -ForegroundColor White
Write-Host "  UTC Offset: $($localZone.BaseUtcOffset)" -ForegroundColor White
Write-Host "  UTC Now: $($utcNow.ToString('HH:mm:ss'))" -ForegroundColor Yellow
Write-Host "  Local Now: $($localNow.ToString('HH:mm:ss'))" -ForegroundColor Green
Write-Host "  GID logs (UTC) will be converted to Local time" -ForegroundColor Gray
Write-Host ""

Load-D2Data

$config = Load-Config -Path $configPath

if ($config.settings_path) {
    if (Test-Path $config.settings_path) {
        Load-CharacterMapping -SettingsPath $config.settings_path
    }
    else {
    }
}

Initialize-ReportedDeaths -Config $config





try {
    Start-Monitoring -Config $config
}
catch {
    Write-Host "`n❌ Error: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit"
}