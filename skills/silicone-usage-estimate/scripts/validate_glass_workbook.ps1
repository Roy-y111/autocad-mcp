param([string]$WorkbookPath=([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\outputs\門窗數量周長估算表_框加玻璃.xlsx'))))
$ErrorActionPreference='Stop';$excel=$null;$book=$null
try{
  $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
  $book=$excel.Workbooks.Open($WorkbookPath);$excel.CalculateFullRebuild()
  $names=@($book.Worksheets|ForEach-Object{$_.Name})
  $required=@('玻璃分割明細','框加玻璃彙總');$missing=@($required|Where-Object{$_-notin$names})
  $formulaCount=0;$errors=@()
  foreach($sheet in @($book.Worksheets)){
    $used=$sheet.UsedRange
    foreach($cell in @($used.Cells)){
      if($cell.HasFormula){$formulaCount++}
      $txt=[string]$cell.Text
      if($txt -match '^#(VALUE!|REF!|NAME\?|DIV/0!|N/A|NUM!|NULL!)$'){$errors+="$($sheet.Name)!$($cell.Address()):$txt"}
    }
  }
  $book.Save();$book.Close($true);$excel.Quit()
  [pscustomobject]@{Sheets=$names.Count;MissingSheets=$missing;FormulaCount=$formulaCount;FormulaErrors=$errors.Count;Errors=$errors}|ConvertTo-Json -Depth 4
  if($missing.Count-or$errors.Count){exit 2}
}finally{if($book){try{$book.Close($false)}catch{}};if($excel){try{$excel.Quit()}catch{}}}
