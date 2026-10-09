param(
 [string]$Source=([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\門窗數量周長估算表.xlsx'))),
 [string]$ResultCsv=([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\silicone_batch_result.csv'))),
 [string]$SeedCsv=([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\silicone_seeds.csv'))),
 [string]$Output=([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\門窗數量周長估算表_聚合線核對版.xlsx')))
)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Output){throw "Output already exists: $Output"}
Copy-Item -LiteralPath $Source -Destination $Output
$result=Import-Csv -LiteralPath $ResultCsv;$seeds=Import-Csv -LiteralPath $SeedCsv
$verified=@{};foreach($g in @($result|Where-Object Status -eq 'OK'|Group-Object Code)){$verified[$g.Name]=[math]::Round(($g.Group|Measure-Object PerimeterCm -Sum).Sum,2)}
$candidateCodes=@{};foreach($c in $seeds.Code){$candidateCodes[$c]=$true}
$excel=$null;$book=$null
try{
 Set-Content -LiteralPath ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\build_verified_progress.txt'))) -Value 'start' -Encoding utf8
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
 $book=$excel.Workbooks.Open($Output)
 Add-Content -LiteralPath ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\build_verified_progress.txt'))) -Value 'opened' -Encoding utf8
 foreach($ws in @($book.Worksheets)){
   Add-Content -LiteralPath ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\build_verified_progress.txt'))) -Value ("sheet="+$ws.Name) -Encoding utf8
   # Detail-floor sheets use row 5 header / row 6 data; all-floor summary uses row 4 header / row 5 data.
   $isSummary=($ws.Name-eq'全樓層彙總');$headerRow=if($isSummary){4}else{5};$row=if($isSummary){5}else{6}
   $lastCol=$ws.UsedRange.Columns.Count+1
   $headers=@('聚合線單面周長(cm)','面數(可調)','單樘玻璃打設長度(cm)','本頁玻璃總長度(cm)','聚合線核對狀態')
   for($i=0;$i-lt$headers.Count;$i++){$ws.Cells.Item($headerRow,$lastCol+$i)=$headers[$i]}
   $hr=$ws.Range($ws.Cells.Item($headerRow,$lastCol),$ws.Cells.Item($headerRow,$lastCol+4));$hr.Font.Bold=$true;$hr.Interior.Color=0x794E1F;$hr.Font.Color=0xFFFFFF;$hr.WrapText=$true
   while($ws.Cells.Item($row,2).Value2){
     $code=[string]$ws.Cells.Item($row,2).Value2
     if($verified.ContainsKey($code)){$ws.Cells.Item($row,$lastCol)=$verified[$code];$ws.Cells.Item($row,$lastCol).Interior.Color=0xDAF2D0;$status='BPOLY＋尺寸校驗通過（紅線）'}
     else{$ws.Cells.Item($row,$lastCol).Interior.Color=0xCCFFFF;$status=if($candidateCodes.ContainsKey($code)){'待核對黃色候選聚合線；確認後手動填周長'}else{'未生成候選聚合線，需人工執行 SIL'}}
     $ws.Cells.Item($row,$lastCol+1)=2;$ws.Cells.Item($row,$lastCol+1).Interior.Color=0xCCFFFF
     $l1=$ws.Cells.Item($row,$lastCol).Address($false,$false);$l2=$ws.Cells.Item($row,$lastCol+1).Address($false,$false)
     $ws.Cells.Item($row,$lastCol+2).Formula="=IF($l1="""","""",$l1*$l2)"
     $qty=$ws.Cells.Item($row,5).Address($false,$false);$two=$ws.Cells.Item($row,$lastCol+2).Address($false,$false)
     $ws.Cells.Item($row,$lastCol+3).Formula="=IF($two="""","""",$two*$qty)"
     $ws.Cells.Item($row,$lastCol+4)=$status
     if($status-like'待核對*'-or$status-like'未生成*'){$ws.Cells.Item($row,$lastCol+4).Interior.Color=0xCCCCFF}else{$ws.Cells.Item($row,$lastCol+4).Interior.Color=0xDAF2D0}
     $row++
   }
   $end=$row-1;$firstDataRow=if($isSummary){5}else{6}
   if($end-ge$firstDataRow){$ws.Range($ws.Cells.Item($headerRow,$lastCol),$ws.Cells.Item($end,$lastCol+4)).Borders.LineStyle=1}
   $ws.Columns.Item($lastCol).ColumnWidth=18;$ws.Columns.Item($lastCol+1).ColumnWidth=11;$ws.Columns.Item($lastCol+2).ColumnWidth=20;$ws.Columns.Item($lastCol+3).ColumnWidth=19;$ws.Columns.Item($lastCol+4).ColumnWidth=38
 }
 Add-Content -LiteralPath ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\build_verified_progress.txt'))) -Value 'sheets-done' -Encoding utf8
 $book.Save();$book.Close($true);$excel.Quit()
 Add-Content -LiteralPath ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\build_verified_progress.txt'))) -Value 'saved' -Encoding utf8
 [pscustomobject]@{Saved=$Output;VerifiedCodes=@($verified.Keys);VerifiedPerimeters=$verified}|ConvertTo-Json -Depth 5
}finally{if($book){try{$book.Close($false)}catch{}};if($excel){try{$excel.Quit()}catch{}}}
