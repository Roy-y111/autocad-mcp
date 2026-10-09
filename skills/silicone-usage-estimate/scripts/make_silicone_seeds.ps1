param(
  [string]$InputPath=([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\silicone_source.txt'))),
  [string]$OutputPath=([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\silicone_seeds.csv')))
)
$lines=Get-Content -LiteralPath $InputPath
$tags=@();$segments=@()
foreach($line in $lines){
  if($line.StartsWith('TAG|')){$p=$line.Split('|');$tags+=[pscustomobject]@{Code=$p[1];X=[double]$p[2];Y=[double]$p[3]}}
  elseif($line.StartsWith('PL|')){
    $p=$line.Split('|');$n=@($p[2].Split(',')|%{[double]$_});$pts=@()
    for($i=0;$i-lt$n.Count;$i+=2){$pts+=,[pscustomobject]@{X=$n[$i];Y=$n[$i+1]}}
    $limit=if($p[1]-eq'1'){$pts.Count}else{$pts.Count-1}
    for($i=0;$i-lt$limit;$i++){$a=$pts[$i];$b=$pts[($i+1)%$pts.Count]
      if([math]::Abs($a.X-$b.X)-lt0.2-and[math]::Abs($a.Y-$b.Y)-gt5){$segments+=[pscustomobject]@{T='V';K=($a.X+$b.X)/2;A=[math]::Min($a.Y,$b.Y);B=[math]::Max($a.Y,$b.Y)}}
      elseif([math]::Abs($a.Y-$b.Y)-lt0.2-and[math]::Abs($a.X-$b.X)-gt5){$segments+=[pscustomobject]@{T='H';K=($a.Y+$b.Y)/2;A=[math]::Min($a.X,$b.X);B=[math]::Max($a.X,$b.X)}}
    }
  }
}
$rowYs=@();foreach($y in @($tags.Y|sort -Descending)){if(!$rowYs.Count-or[math]::Abs($rowYs[-1]-$y)-gt40){$rowYs+=$y}}
$bounds=@{}
for($r=0;$r-lt$rowYs.Count;$r++){$ry=$rowYs[$r];$row=@($tags|?{[math]::Abs($_.Y-$ry)-le40}|sort X);$ny=if($r+1-lt$rowYs.Count){$rowYs[$r+1]}else{$null}
  for($i=0;$i-lt$row.Count;$i++){$t=$row[$i];$bounds[$t.Code]=[pscustomobject]@{X0=$t.X+50;X1=if($i+1-lt$row.Count){$row[$i+1].X+50}else{$t.X+600};Y0=if($null-ne$ny){$ny+60}else{$t.Y-850};Y1=$t.Y+60}}
}
function UniqueCoord($vals){$out=@();foreach($v in @($vals|sort)){if(!$out.Count-or[math]::Abs($out[-1]-$v)-gt0.5){$out+=$v}};$out}
function Covers($list,$k,$a,$b){@($list|?{[math]::Abs($_.K-$k)-le0.5-and$_.A-le($a+0.5)-and$_.B-ge($b-0.5)}).Count-gt0}
$out=@()
foreach($t in @($tags|?{$_.Code-match'^(W|DW|WB)' }|sort @{e='Y';Descending=$true},X)){$b=$bounds[$t.Code]
  $vs=@($segments|?{$_.T-eq'V'-and$b.X0-le$_.K-and$_.K-le$b.X1-and$_.B-ge$b.Y0-and$_.A-le$b.Y1})
  $hs=@($segments|?{$_.T-eq'H'-and$b.Y0-le$_.K-and$_.K-le$b.Y1-and$_.B-ge$b.X0-and$_.A-le$b.X1})
  $xs=@(UniqueCoord $vs.K);$ys=@(UniqueCoord $hs.K);$idx=0
  if($t.Code-eq'W1'){Write-Output "W1 bounds=$($b.X0),$($b.X1),$($b.Y0),$($b.Y1) V=$($vs.Count) H=$($hs.Count) X=$($xs.Count) Y=$($ys.Count)";Write-Output ('Xs='+($xs-join','));Write-Output ('Ys='+($ys-join','))}
  for($xi=0;$xi+1-lt$xs.Count;$xi++){for($yi=0;$yi+1-lt$ys.Count;$yi++){$x0=$xs[$xi];$x1=$xs[$xi+1];$y0=$ys[$yi];$y1=$ys[$yi+1];$w=$x1-$x0;$h=$y1-$y0
    if($w-gt10-and$h-gt10-and(Covers $vs $x0 $y0 $y1)-and(Covers $vs $x1 $y0 $y1)-and(Covers $hs $y0 $x0 $x1)-and(Covers $hs $y1 $x0 $x1)){$idx++;$out+=[pscustomobject]@{Code=$t.Code;Pane="G$idx";X=[math]::Round(($x0+$x1)/2,4);Y=[math]::Round(($y0+$y1)/2,4);ExpectedW=[math]::Round($w,2);ExpectedH=[math]::Round($h,2)}}
  }}
}
$out|Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
$pipePath=[IO.Path]::ChangeExtension($OutputPath,'.txt')
@($out|ForEach-Object{"$($_.Code)|$($_.Pane)|$($_.X)|$($_.Y)|$($_.ExpectedW)|$($_.ExpectedH)"})|Set-Content -LiteralPath $pipePath -Encoding ascii
$out|group Code|select Name,Count|sort Name|ft -AutoSize
Write-Output "Saved: $OutputPath"
