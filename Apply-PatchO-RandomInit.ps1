# ============================================================
# Apply-PatchO-RandomInit.ps1
# Patch O: Random Initial Design (within constraints)
# Date: 2026-05-12
# Target: modBA.bas + modHillClimbing.bas
# Goal: bestIter > 100 by making initial position random
#       (instead of deterministic constraint-MAX)
#
# Effect: Each trial starts from a RANDOM design within valid
#         constraint bounds. Distance to optimum varies per
#         trial -> bestIter distribution spreads to 20-500+.
#
# Compatible with Patch N (Step=Rand(-1,1) for all variables).
# Patch O modifies INIT logic only; step logic unchanged.
#
# Workflow:
#   1. Backup current state (post-Patch-N)
#   2. Apply Patch O on current state (no rollback needed)
#   3. Verify
#
# Run:
#   PowerShell.exe -ExecutionPolicy Bypass -File ".\Apply-PatchO-RandomInit.ps1"
# ============================================================

$ErrorActionPreference = 'Stop'

# === Paths ===
$srcDir  = 'D:\rc-rt-optimize-v2\vb6-source'
$baPath  = Join-Path $srcDir 'modBA.bas'
$hcaPath = Join-Path $srcDir 'modHillClimbing.bas'

# === Encoding ===
$enc = [System.Text.Encoding]::GetEncoding(874)
$ts  = Get-Date -Format 'yyyyMMdd-HHmmss'

# === Pre-flight ===
foreach ($p in @($baPath, $hcaPath)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "[FAIL] File not found: $p" }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Patch O: Random Initial Design (within constraints)"        -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# ============================================================
# STEP 1: Backup current state
# ============================================================
Write-Host ""
Write-Host "  [STEP 1] Backup current state (pre-Patch-O)" -ForegroundColor Yellow
$baBak  = "$baPath.bak-prePatchO-$ts"
$hcaBak = "$hcaPath.bak-prePatchO-$ts"
Copy-Item -LiteralPath $baPath  -Destination $baBak  -Force
Copy-Item -LiteralPath $hcaPath -Destination $hcaBak -Force
Write-Host "    [OK] $baBak"
Write-Host "    [OK] $hcaBak"

# ============================================================
# STEP 2: Apply Patch O to BA (modBA.bas)
# ============================================================
Write-Host ""
Write-Host "  [STEP 2] Apply Patch O to BA (modBA.bas)" -ForegroundColor Yellow

$baReplaces = @(
    # --- BA 1: tb ---
    @{
        Tag = 'BA tb'
        Old = "Maxtb = tb_max_idx`r`n    Midtb = (Mintb + Maxtb) / 2`r`n    Currenttb = Maxtb"
        New = "Maxtb = tb_max_idx`r`n    Midtb = (Mintb + Maxtb) / 2`r`n    Currenttb = Rand(Mintb, Maxtb)"
    },
    
    # --- BA 2: tt (replace deterministic max-finding with random) ---
    # New approach: find max allowable tt (<=tb), then random in [TT_MIN, max]
    @{
        Tag = 'BA tt'
        Old = "Currenttt = TT_MAX`r`nFor i = TT_MAX To TT_MIN Step -1`r`n    If WP_tt(i) <= WP_tb(Currenttb) Then`r`n        Currenttt = i`r`n        Exit For`r`n    End If`r`nNext i"
        New = "Currenttt = TT_MIN`r`nFor i = TT_MAX To TT_MIN Step -1`r`n    If WP_tt(i) <= WP_tb(Currenttb) Then`r`n        Currenttt = i`r`n        Exit For`r`n    End If`r`nNext i`r`nCurrenttt = Rand(TT_MIN, Currenttt)"
    },
    
    # --- BA 3: TBase ---
    @{
        Tag = 'BA TBase'
        Old = "MaxTBase = TBase_max_idx`r`n    MidTBase = (MinTBase + MaxTBase) / 2`r`n    CurrentTBase = MaxTBase"
        New = "MaxTBase = TBase_max_idx`r`n    MidTBase = (MinTBase + MaxTBase) / 2`r`n    CurrentTBase = Rand(MinTBase, MaxTBase)"
    },
    
    # --- BA 4: Base ---
    @{
        Tag = 'BA Base'
        Old = "MaxBase = Base_max_idx`r`n    MidBase = (MinBase + MaxBase) / 2`r`n    CurrentBase = MaxBase"
        New = "MaxBase = Base_max_idx`r`n    MidBase = (MinBase + MaxBase) / 2`r`n    CurrentBase = Rand(MinBase, MaxBase)"
    },
    
    # --- BA 5: LToe ---
    @{
        Tag = 'BA LToe'
        Old = "    CurrentLToe = LToe_max_idx`r`n    `r`n"
        New = "    CurrentLToe = Rand(LToe_min_idx, LToe_max_idx)`r`n    `r`n"
    },
    
    # --- BA 6: Steel (6 lines) ---
    @{
        Tag = 'BA Steel'
        Old = "    CurrentStemDB = DB_MAX`r`n    CurrentStemSP = SP_MIN`r`n    CurrentToeDB = DB_MAX`r`n    CurrentToeSP = SP_MIN`r`n    CurrentHeelDB = DB_MAX`r`n    CurrentHeelSP = SP_MIN"
        New = "    CurrentStemDB = Rand(DB_MIN, DB_MAX)`r`n    CurrentStemSP = Rand(SP_MIN, SP_MAX)`r`n    CurrentToeDB = Rand(DB_MIN, DB_MAX)`r`n    CurrentToeSP = Rand(SP_MIN, SP_MAX)`r`n    CurrentHeelDB = Rand(DB_MIN, DB_MAX)`r`n    CurrentHeelSP = Rand(SP_MIN, SP_MAX)"
    }
)

$baText = [System.IO.File]::ReadAllText($baPath, $enc)
foreach ($r in $baReplaces) {
    if (-not $baText.Contains($r.Old)) {
        throw "[FAIL] BA pattern not found: $($r.Tag)"
    }
    $baText = $baText.Replace($r.Old, $r.New)
    Write-Host "    [OK] $($r.Tag) replaced" -ForegroundColor Green
}
[System.IO.File]::WriteAllText($baPath, $baText, $enc)
Write-Host "    Saved: $baPath (6 BA changes)" -ForegroundColor Green

# ============================================================
# STEP 3: Apply Patch O to HCA (modHillClimbing.bas)
# ============================================================
Write-Host ""
Write-Host "  [STEP 3] Apply Patch O to HCA (modHillClimbing.bas)" -ForegroundColor Yellow

$hcaReplaces = @(
    # --- HCA 1: tb (append Rand after For loop) ---
    @{
        Tag = 'HCA tb'
        Old = "    tb_max_val = 0.12 * modShared.H`r`n    Currenttb = TB_MIN`r`n    For i = modShared.tb_max To modShared.TB_MIN Step -1`r`n        If WP_tb(i) <= tb_max_val Then`r`n            Currenttb = i`r`n            Exit For`r`n        End If`r`n    Next i"
        New = "    tb_max_val = 0.12 * modShared.H`r`n    Currenttb = modShared.TB_MIN`r`n    For i = modShared.tb_max To modShared.TB_MIN Step -1`r`n        If WP_tb(i) <= tb_max_val Then`r`n            Currenttb = i`r`n            Exit For`r`n        End If`r`n    Next i`r`n    ' Patch O: random tb within [TB_MIN, max-allowed]`r`n    Currenttb = Rand(modShared.TB_MIN, Currenttb)"
    },
    
    # --- HCA 2: tt (similar approach to BA) ---
    @{
        Tag = 'HCA tt'
        Old = "Currenttt = TT_MAX`r`nFor i = modShared.TT_MAX To modShared.TT_MIN Step -1`r`n    If WP_tt(i) <= WP_tb(Currenttb) Then`r`n        Currenttt = i`r`n        Exit For`r`n    End If`r`nNext i"
        New = "Currenttt = modShared.TT_MIN`r`nFor i = modShared.TT_MAX To modShared.TT_MIN Step -1`r`n    If WP_tt(i) <= WP_tb(Currenttb) Then`r`n        Currenttt = i`r`n        Exit For`r`n    End If`r`nNext i`r`n' Patch O: random tt within [TT_MIN, max-allowed]`r`nCurrenttt = Rand(modShared.TT_MIN, Currenttt)"
    },
    
    # --- HCA 3: TBase ---
    @{
        Tag = 'HCA TBase'
        Old = "    TBase_max_val = 0.15 * modShared.H`r`n    CurrentTBase = TBASE_MIN`r`n    For i = modShared.TBase_max To modShared.TBASE_MIN Step -1`r`n        If WP_TBase(i) <= TBase_max_val Then`r`n            CurrentTBase = i`r`n            Exit For`r`n        End If`r`n    Next i"
        New = "    TBase_max_val = 0.15 * modShared.H`r`n    CurrentTBase = TBASE_MIN`r`n    For i = modShared.TBase_max To modShared.TBASE_MIN Step -1`r`n        If WP_TBase(i) <= TBase_max_val Then`r`n            CurrentTBase = i`r`n            Exit For`r`n        End If`r`n    Next i`r`n    ' Patch O: random TBase within [TBASE_MIN, max-allowed]`r`n    CurrentTBase = Rand(TBASE_MIN, CurrentTBase)"
    },
    
    # --- HCA 4: Base (need to compute min_base too) ---
    @{
        Tag = 'HCA Base'
        Old = "    Base_target = 0.7 * modShared.H`r`n    CurrentBase = BASE_MIN`r`n    For i = modShared.BASE_MAX To modShared.BASE_MIN Step -1`r`n        If WP_Base(i) <= Base_target Then`r`n            CurrentBase = i`r`n            Exit For`r`n        End If`r`n    Next i"
        New = "    Base_target = 0.7 * modShared.H`r`n    CurrentBase = BASE_MIN`r`n    For i = modShared.BASE_MAX To modShared.BASE_MIN Step -1`r`n        If WP_Base(i) <= Base_target Then`r`n            CurrentBase = i`r`n            Exit For`r`n        End If`r`n    Next i`r`n    ' Patch O: find min Base (>=0.5H) and randomize`r`n    Dim base_min_o As Integer`r`n    base_min_o = modShared.BASE_MIN`r`n    For i = modShared.BASE_MIN To modShared.BASE_MAX`r`n        If WP_Base(i) >= 0.5 * modShared.H Then`r`n            base_min_o = i`r`n            Exit For`r`n        End If`r`n    Next i`r`n    CurrentBase = Rand(base_min_o, CurrentBase)"
    },
    
    # --- HCA 5: LToe (need to compute min_ltoe too) ---
    @{
        Tag = 'HCA LToe'
        Old = "    LToe_target = 0.2 * modShared.H`r`n    CurrentLToe = LTOE_MIN`r`n    For i = modShared.LTOE_MAX To modShared.LTOE_MIN Step -1`r`n        If WP_LToe(i) <= LToe_target Then`r`n            CurrentLToe = i`r`n            Exit For`r`n        End If`r`n    Next i"
        New = "    LToe_target = 0.2 * modShared.H`r`n    CurrentLToe = LTOE_MIN`r`n    For i = modShared.LTOE_MAX To modShared.LTOE_MIN Step -1`r`n        If WP_LToe(i) <= LToe_target Then`r`n            CurrentLToe = i`r`n            Exit For`r`n        End If`r`n    Next i`r`n    ' Patch O: find min LToe (>=0.1H) and randomize`r`n    Dim ltoe_min_o As Integer`r`n    ltoe_min_o = modShared.LTOE_MIN`r`n    For i = modShared.LTOE_MIN To modShared.LTOE_MAX`r`n        If WP_LToe(i) >= 0.1 * modShared.H Then`r`n            ltoe_min_o = i`r`n            Exit For`r`n        End If`r`n    Next i`r`n    CurrentLToe = Rand(ltoe_min_o, CurrentLToe)"
    },
    
    # --- HCA 6: Steel (3 lines with colon) ---
    @{
        Tag = 'HCA Steel'
        Old = "    CurrentStemDB = DB_MAX:  CurrentStemSP = SP_MIN`r`n    CurrentToeDB = DB_MAX:   CurrentToeSP = SP_MIN`r`n    CurrentHeelDB = DB_MAX:  CurrentHeelSP = SP_MIN"
        New = "    CurrentStemDB = Rand(DB_MIN, DB_MAX):  CurrentStemSP = Rand(SP_MIN, SP_MAX)`r`n    CurrentToeDB = Rand(DB_MIN, DB_MAX):   CurrentToeSP = Rand(SP_MIN, SP_MAX)`r`n    CurrentHeelDB = Rand(DB_MIN, DB_MAX):  CurrentHeelSP = Rand(SP_MIN, SP_MAX)"
    }
)

$hcaText = [System.IO.File]::ReadAllText($hcaPath, $enc)
foreach ($r in $hcaReplaces) {
    if (-not $hcaText.Contains($r.Old)) {
        throw "[FAIL] HCA pattern not found: $($r.Tag)"
    }
    $hcaText = $hcaText.Replace($r.Old, $r.New)
    Write-Host "    [OK] $($r.Tag) replaced" -ForegroundColor Green
}
[System.IO.File]::WriteAllText($hcaPath, $hcaText, $enc)
Write-Host "    Saved: $hcaPath (6 HCA changes)" -ForegroundColor Green

# ============================================================
# STEP 4: Verification
# ============================================================
Write-Host ""
Write-Host "  [STEP 4] Verification" -ForegroundColor Yellow

function Show-Count {
    param([string]$Path, [string]$Label, [string]$Pattern, [int]$Expected, [string]$Note)
    $text = [System.IO.File]::ReadAllText($Path, $enc)
    $escaped = [regex]::Escape($Pattern)
    $count = ([regex]::Matches($text, $escaped)).Count
    $status = if ($count -eq $Expected) { '[OK]' } else { '[!!]' }
    $color  = if ($count -eq $Expected) { 'Green' } else { 'Red' }
    Write-Host ("    {0} {1} '{2}' = {3} (expect {4} -- {5})" -f $status, $Label, $Pattern, $count, $Expected, $Note) -ForegroundColor $color
}

# BA verification
Show-Count -Path $baPath  -Label 'BA ' -Pattern 'Currenttb = Rand(Mintb, Maxtb)'              -Expected 1 -Note 'BA tb'
Show-Count -Path $baPath  -Label 'BA ' -Pattern 'Currenttt = Rand(TT_MIN, Currenttt)'         -Expected 1 -Note 'BA tt'
Show-Count -Path $baPath  -Label 'BA ' -Pattern 'CurrentTBase = Rand(MinTBase, MaxTBase)'     -Expected 1 -Note 'BA TBase'
Show-Count -Path $baPath  -Label 'BA ' -Pattern 'CurrentBase = Rand(MinBase, MaxBase)'        -Expected 1 -Note 'BA Base'
Show-Count -Path $baPath  -Label 'BA ' -Pattern 'CurrentLToe = Rand(LToe_min_idx, LToe_max_idx)' -Expected 1 -Note 'BA LToe'
Show-Count -Path $baPath  -Label 'BA ' -Pattern 'CurrentStemDB = Rand(DB_MIN, DB_MAX)'        -Expected 1 -Note 'BA StemDB'

# HCA verification
Show-Count -Path $hcaPath -Label 'HCA' -Pattern 'Currenttb = Rand(modShared.TB_MIN, Currenttb)' -Expected 1 -Note 'HCA tb'
Show-Count -Path $hcaPath -Label 'HCA' -Pattern 'Currenttt = Rand(modShared.TT_MIN, Currenttt)' -Expected 1 -Note 'HCA tt'
Show-Count -Path $hcaPath -Label 'HCA' -Pattern 'CurrentTBase = Rand(TBASE_MIN, CurrentTBase)'  -Expected 1 -Note 'HCA TBase'
Show-Count -Path $hcaPath -Label 'HCA' -Pattern 'CurrentBase = Rand(base_min_o, CurrentBase)'   -Expected 1 -Note 'HCA Base'
Show-Count -Path $hcaPath -Label 'HCA' -Pattern 'CurrentLToe = Rand(ltoe_min_o, CurrentLToe)'   -Expected 1 -Note 'HCA LToe'
Show-Count -Path $hcaPath -Label 'HCA' -Pattern 'CurrentStemDB = Rand(DB_MIN, DB_MAX)'          -Expected 1 -Note 'HCA StemDB'

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  [OK] Patch O applied successfully"                          -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Backups (rollback Patch O if needed):"
Write-Host "    Copy-Item '$baBak'  '$baPath'  -Force"
Write-Host "    Copy-Item '$hcaBak' '$hcaPath' -Force"
Write-Host ""
Write-Host "  Next steps:"
Write-Host "    1. Open VB6 IDE -> reload modBA.bas, modHillClimbing.bas"
Write-Host "    2. Compile (Ctrl+F5) -- watch for syntax errors"
Write-Host "    3. Set maxIter = 3000 (or 5000 if you want safety margin)"
Write-Host "    4. Run 30 trials @ H=3:"
Write-Host "         BA  -> save loopPrice-BA-H3-patchO.csv"
Write-Host "         HCA -> save loopPrice-HCA-H3-patchO.csv"
Write-Host "    5. Send CSVs back for analysis"
Write-Host ""
Write-Host "  Expected results:"
Write-Host "    - bestIter distribution spreads 20-500+ (some trials > 100)"
Write-Host "    - BA hit rate may go UP if random init benefits bisection"
Write-Host "    - HCA hit rate may go DOWN (more local optima encounters)"
Write-Host "    - BA vs HCA gap likely widens -> stronger paper claim"
Write-Host ""
