param([string]$Path=([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\門窗數量周長估算表_聚合線核對版.xlsx'))))
$ErrorActionPreference='Stop';$excel=$null;$book=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
 $book=$excel.Workbooks.Open($Path);$excel.CalculateFullRebuild();$book.Save()
 $errors=@();$formulaCount=0
 foreach($ws in @($book.Worksheets)){foreach($cell in @($ws.UsedRange.Cells)){if($cell.HasFormula){$formulaCount++};$txt=[string]$cell.Text;if($txt-match'^#(VALUE!|REF!|NAME\?|DIV/0!|N/A|NUM!|NULL!)$'){$errors+="$($ws.Name)!$($cell.Address()):$txt"}}}
 $sum=$book.Worksheets.Item('全樓層彙總');$lastCol=$sum.UsedRange.Columns.Count
 $headers=@{};for($c=1;$c-le$lastCol;$c++){$headers[[string]$sum.Cells.Item(4,$c).Value2]=$c}
 $checks=@{};$r=5
 while($sum.Cells.Item($r,2).Value2){$code=[string]$sum.Cells.Item($r,2).Value2;if($code-in@('W1','W2')){$checks[$code]=[pscustomobject]@{Perimeter=$sum.Cells.Item($r,$headers['聚合線單面周長(cm)']).Value2;Faces=$sum.Cells.Item($r,$headers['面數(可調)']).Value2;Status=$sum.Cells.Item($r,$headers['聚合線核對狀態']).Value2}};$r++}
 $book.Close($true);$excel.Quit()
 [pscustomobject]@{FormulaCount=$formulaCount;FormulaErrors=$errors.Count;Errors=$errors;Checks=$checks}|ConvertTo-Json -Depth 5
 if($errors.Count-or$checks['W1'].Perimeter-ne1280-or$checks['W2'].Perimeter-ne602-or$checks['W1'].Faces-ne2-or$checks['W2'].Faces-ne2){exit 2}
}finally{if($book){try{$book.Close($false)}catch{}};if($excel){try{$excel.Quit()}catch{}}}
