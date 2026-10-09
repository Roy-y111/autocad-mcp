param(
    [string]$InputPath = ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\a6_glass_raw.txt'))),
    [string]$OutputPath = ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\a6_glass_analysis.json')))
)

$raw = Get-Content -LiteralPath $InputPath
$tags = @()
$fixes = @()
$polys = @()

for ($i = 0; $i -lt $raw.Count; $i++) {
    $line = $raw[$i]
    if ($line.StartsWith('TAG|')) {
        $p = $line.Split('|')
        $tags += [pscustomobject]@{ Block=$p[1]; Code=$p[2]; X=[double]$p[3]; Y=[double]$p[4] }
    } elseif ($line.StartsWith('FIX|')) {
        $p = $line.Split('|')
        $fixes += [pscustomobject]@{ X=[double]$p[1]; Y=[double]$p[2] }
    } elseif ($line.StartsWith('PL|')) {
        $p = $line.Split('|')
        if ($i + 1 -ge $raw.Count) { continue }
        $nums = @($raw[++$i].Split(',') | ForEach-Object { [double]$_ })
        if ($nums.Count -lt 4) { continue }
        $xs = @(); $ys = @()
        for ($j=0; $j -lt $nums.Count; $j+=2) { $xs += $nums[$j]; $ys += $nums[$j+1] }
        $polys += [pscustomobject]@{
            Layer=$p[1]; Closed=($p[2] -eq '1')
            X0=($xs|Measure-Object -Minimum).Minimum; X1=($xs|Measure-Object -Maximum).Maximum
            Y0=($ys|Measure-Object -Minimum).Minimum; Y1=($ys|Measure-Object -Maximum).Maximum
            PointCount=$xs.Count
        }
    }
}

$windowTags = @($tags | Where-Object { $_.Code -match '^(W|DW|WB)' })
$rowYs = @()
foreach ($y in @($tags.Y | Sort-Object -Descending)) {
    if ($rowYs.Count -eq 0 -or [math]::Abs($rowYs[-1] - $y) -gt 40) { $rowYs += $y }
}

$bounds = @{}
for ($ri=0; $ri -lt $rowYs.Count; $ri++) {
    $ry = $rowYs[$ri]
    $row = @($tags | Where-Object { [math]::Abs($_.Y-$ry) -le 40 } | Sort-Object X)
    $nextY = if ($ri+1 -lt $rowYs.Count) { $rowYs[$ri+1] } else { $null }
    for ($k=0; $k -lt $row.Count; $k++) {
        $t=$row[$k]
        $x0=if($k -gt 0){($row[$k-1].X+$t.X)/2}else{$t.X-300}
        $x1=if($k+1 -lt $row.Count){($row[$k+1].X+$t.X)/2}else{$t.X+300}
        $y1=$t.Y+60
        $y0=if($null -ne $nextY){$nextY+60}else{$t.Y-850}
        $bounds[$t.Code]=[pscustomobject]@{X0=$x0;X1=$x1;Y0=$y0;Y1=$y1}
    }
}

function Contains-Rect($outer,$inner) {
    return ($outer.X0-2 -le $inner.X0 -and $inner.X1 -le $outer.X1+2 -and
            $outer.Y0-2 -le $inner.Y0 -and $inner.Y1 -le $outer.Y1+2)
}

$result = @()
foreach ($t in @($windowTags | Sort-Object @{Expression='Y';Descending=$true}, @{Expression='X';Descending=$false})) {
    $b=$bounds[$t.Code]
    $rects=@($polys | Where-Object {
        $_.Layer -eq 'OPEN' -and $_.Closed -and ($_.X1-$_.X0) -gt 5 -and ($_.Y1-$_.Y0) -gt 5 -and
        $b.X0 -le $_.X0 -and $_.X1 -le $b.X1 -and $b.Y0 -le $_.Y0 -and $_.Y1 -le $b.Y1
    })
    $leaf=@()
    foreach($r in $rects){
        $isFrame=$false
        $area=($r.X1-$r.X0)*($r.Y1-$r.Y0)
        foreach($r2 in $rects){
            if([object]::ReferenceEquals($r,$r2)){continue}
            $area2=($r2.X1-$r2.X0)*($r2.Y1-$r2.Y0)
            if((Contains-Rect $r $r2) -and $area2 -lt $area*0.95){$isFrame=$true;break}
        }
        if(-not $isFrame){$leaf += $r}
    }
    $glass=@()
    $n=0
    foreach($r in @($leaf|Sort-Object X0,Y0)){
        $n++
        $w=[math]::Round($r.X1-$r.X0,2);$h=[math]::Round($r.Y1-$r.Y0,2)
        $glass += [pscustomobject]@{Id="G$n";Width=$w;Height=$h;Perimeter=[math]::Round(2*($w+$h),2)}
    }
    $fixCount=@($fixes|Where-Object{$b.X0-le$_.X-and$_.X-le$b.X1-and$b.Y0-le$_.Y-and$_.Y-le$b.Y1}).Count
    $marks=@($polys|Where-Object{$_.Layer-eq'C7-0'-and$b.X0-le(($_.X0+$_.X1)/2)-and(($_.X0+$_.X1)/2)-le$b.X1-and$b.Y0-le(($_.Y0+$_.Y1)/2)-and(($_.Y0+$_.Y1)/2)-le$b.Y1}).Count
    $status=if($leaf.Count -eq 0){'無法辨識'}elseif(($fixCount+$marks) -gt $leaf.Count){'分割推估'}else{'幾何辨識'}
    $result += [pscustomobject]@{Code=$t.Code;Status=$status;GlassCandidateCount=$leaf.Count;FixCount=$fixCount;OpenMarkCount=$marks;Glass=$glass}
}

$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding utf8
$result | Select-Object Code,Status,GlassCandidateCount,FixCount,OpenMarkCount | Export-Csv -LiteralPath ([IO.Path]::ChangeExtension($OutputPath,'.csv')) -NoTypeInformation -Encoding utf8
$result | Group-Object Status | Select-Object Name,Count | Format-Table -AutoSize
Write-Output "Saved: $OutputPath"
