param(
    [string]$SourceWorkbook = ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\門窗數量周長估算表.xlsx'))),
    [string]$AnalysisJson = ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\a6_glass_analysis.json'))),
    [string]$OutputWorkbook = ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\門窗數量周長估算表_框加玻璃.xlsx')))
)

$ErrorActionPreference = 'Stop'
if (Test-Path -LiteralPath $OutputWorkbook) { throw "Output already exists: $OutputWorkbook" }
Copy-Item -LiteralPath $SourceWorkbook -Destination $OutputWorkbook
$analysis = @(Get-Content -LiteralPath $AnalysisJson -Raw | ConvertFrom-Json)

$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($OutputWorkbook)
    $source = $book.Worksheets.Item('全樓層彙總')

    foreach ($name in @('玻璃分割明細','框加玻璃彙總')) {
        foreach ($sheet in @($book.Worksheets)) { if ($sheet.Name -eq $name) { $sheet.Delete() } }
    }
    $detail = $book.Worksheets.Add(); $detail.Name = '玻璃分割明細'
    $summary = $book.Worksheets.Add(); $summary.Name = '框加玻璃彙總'

    $blue = 0x794E1F; $lightBlue = 0xF0E4D6; $yellow = 0xCCFFFF; $red = 0xCCCCFF; $green = 0xDAF2D0

    # Detail sheet: reserve 20 pane rows per code so users can fill missing panes without breaking SUMIF ranges.
    $detail.Cells.Item(1,1) = '門窗玻璃分割明細（黃色儲存格可人工校正）'
    $detail.Range('A1:J1').Merge(); $detail.Range('A1:J1').Font.Bold=$true; $detail.Range('A1:J1').Font.Size=14
    $detail.Cells.Item(2,1) = '幾何尺寸取自門窗立面圖 OPEN 封閉矩形，屬估算值，不等同玻璃下料尺寸。每型號預留20列；缺漏可自行補入寬高。'
    $detail.Range('A2:J2').Merge()
    $headers=@('編號','辨識狀態','玻璃格','寬(cm)','高(cm)','單面周長(cm)','玻璃打設面數','單樘本格長度(cm)','資料來源','備註')
    for($c=1;$c-le$headers.Count;$c++){$detail.Cells.Item(4,$c)=$headers[$c-1]}
    $detail.Range('A4:J4').Font.Bold=$true; $detail.Range('A4:J4').Interior.Color=$blue; $detail.Range('A4:J4').Font.Color=0xFFFFFF

    $row=5
    $seen=@{}
    foreach($a in $analysis){
        if($seen.ContainsKey($a.Code)){continue};$seen[$a.Code]=$true
        $paneById=@{};foreach($g in @($a.Glass)){$paneById[$g.Id]=$g}
        for($slot=1;$slot-le20;$slot++){
            $id="G$slot";$g=$paneById[$id]
            $detail.Cells.Item($row,1)=$a.Code; $detail.Cells.Item($row,2)=$a.Status; $detail.Cells.Item($row,3)=$id
            if($null-ne$g){$detail.Cells.Item($row,4)=$g.Width;$detail.Cells.Item($row,5)=$g.Height;$detail.Cells.Item($row,9)='CAD幾何辨識'}
            else{$detail.Cells.Item($row,9)='人工輸入'}
            $detail.Cells.Item($row,6).Formula="=IF(OR(D$row="""",E$row=""""),0,2*(D$row+E$row))"
            $detail.Cells.Item($row,7)=2
            $detail.Cells.Item($row,8).Formula="=F$row*G$row"
            $detail.Range("D$row:E$row").Interior.Color=$yellow; $detail.Cells.Item($row,7).Interior.Color=$yellow
            if($a.Status-ne'幾何辨識'){$detail.Cells.Item($row,2).Interior.Color=$red}
            else{$detail.Cells.Item($row,2).Interior.Color=$green}
            if($slot-gt@($a.Glass).Count){$detail.Range("D$row:E$row").Interior.Color=$yellow}
            $row++
        }
    }
    $detail.Range("A4:J$($row-1)").Borders.LineStyle=1
    $detail.Columns.Item(1).ColumnWidth=10;$detail.Columns.Item(2).ColumnWidth=12;$detail.Columns.Item(3).ColumnWidth=9
    $detail.Columns.Item(4).ColumnWidth=10;$detail.Columns.Item(5).ColumnWidth=10;$detail.Columns.Item(6).ColumnWidth=14
    $detail.Columns.Item(7).ColumnWidth=14;$detail.Columns.Item(8).ColumnWidth=17;$detail.Columns.Item(9).ColumnWidth=14;$detail.Columns.Item(10).ColumnWidth=35
    $detail.Application.ActiveWindow.SplitRow=4;$detail.Application.ActiveWindow.FreezePanes=$true

    # Summary sheet based on the original all-floor schedule.
    $summary.Cells.Item(1,1)='全樓層 門窗框＋玻璃矽利康估算'
    $summary.Range('A1:O1').Merge();$summary.Range('A1:O1').Font.Bold=$true;$summary.Range('A1:O1').Font.Size=14
    $summary.Cells.Item(2,1)='每支(管)可施工長度(m，可調)';$summary.Cells.Item(2,2)=6;$summary.Cells.Item(2,2).Interior.Color=$yellow
    $summary.Cells.Item(2,4)='框打設面數預設1；玻璃打設面數在「玻璃分割明細」逐格設定，預設內外2面。'
    $summary.Range('D2:O2').Merge()
    $h=@('類別','編號','總數量','框單面長度(cm)','框打設面數','框單樘長度(cm)','玻璃辨識狀態','玻璃候選片數','玻璃單面周長合計(cm)','玻璃周長人工校正(cm)','採用玻璃單面周長(cm)','玻璃單樘長度(cm)','框+玻璃單樘(cm)','全案總長(m)','用量(支/管)')
    for($c=1;$c-le$h.Count;$c++){$summary.Cells.Item(4,$c)=$h[$c-1]}
    $summary.Range('A4:O4').Font.Bold=$true;$summary.Range('A4:O4').Interior.Color=$blue;$summary.Range('A4:O4').Font.Color=0xFFFFFF
    $srcRow=5;$outRow=5
    while($source.Cells.Item($srcRow,2).Value2){
        $code=[string]$source.Cells.Item($srcRow,2).Value2
        $a=@($analysis|Where-Object Code -eq $code|Select-Object -First 1)
        $summary.Cells.Item($outRow,1)=$source.Cells.Item($srcRow,1).Value2
        $summary.Cells.Item($outRow,2)=$code
        $summary.Cells.Item($outRow,3)=$source.Cells.Item($srcRow,5).Value2
        $summary.Cells.Item($outRow,4).Formula="='全樓層彙總'!H$srcRow"
        $summary.Cells.Item($outRow,5)=1;$summary.Cells.Item($outRow,5).Interior.Color=$yellow
        $summary.Cells.Item($outRow,6).Formula="=D$outRow*E$outRow"
        if($a.Count){$summary.Cells.Item($outRow,7)=$a[0].Status;$summary.Cells.Item($outRow,8)=$a[0].GlassCandidateCount}else{$summary.Cells.Item($outRow,7)='無玻璃分割資料';$summary.Cells.Item($outRow,8)=0}
        $summary.Cells.Item($outRow,9).Formula="=SUMIF('玻璃分割明細'!A:A,B$outRow,'玻璃分割明細'!F:F)"
        $summary.Cells.Item($outRow,10).Interior.Color=$yellow
        $summary.Cells.Item($outRow,11).Formula="=IF(J$outRow="""",I$outRow,J$outRow)"
        $summary.Cells.Item($outRow,12).Formula="=SUMIF('玻璃分割明細'!A:A,B$outRow,'玻璃分割明細'!H:H)"
        $summary.Cells.Item($outRow,13).Formula="=F$outRow+L$outRow"
        $summary.Cells.Item($outRow,14).Formula="=M$outRow*C$outRow/100"
        $summary.Cells.Item($outRow,15).Formula="=ROUNDUP(N$outRow/`$B`$2,0)"
        if($summary.Cells.Item($outRow,7).Value2-ne'幾何辨識'){$summary.Cells.Item($outRow,7).Interior.Color=$red}else{$summary.Cells.Item($outRow,7).Interior.Color=$green}
        $srcRow++;$outRow++
    }
    $summary.Range("A4:O$($outRow-1)").Borders.LineStyle=1
    $summary.Range("D5:O$($outRow-1)").NumberFormat='0.00'
    for($c=1;$c-le15;$c++){$summary.Columns.Item($c).ColumnWidth=if($c-in10,11,12,13,14){17}elseif($c-eq7){14}else{11}}
    $summary.Application.ActiveWindow.SplitRow=4;$summary.Application.ActiveWindow.FreezePanes=$true

    $book.Save()
    $book.Close($true)
    $excel.Quit()
    Write-Output "Saved: $OutputWorkbook"
}
finally {
    if($book){try{$book.Close($false)}catch{}}
    if($excel){try{$excel.Quit()}catch{}}
    foreach($o in @($source,$detail,$summary,$book,$excel)){if($o){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}catch{}}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

