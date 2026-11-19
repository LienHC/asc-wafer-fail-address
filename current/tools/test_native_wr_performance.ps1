param([string]$Workbook,[string]$Output)
$ErrorActionPreference='Stop';$app=$null;$book=$null
function Check($ok,$message){if(-not $ok){throw $message}}
function Run-Report {
 try{$app.Run("'"+$book.Name+"'!WRPerformance.BuildReport",$book) | Out-Null}
 catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw}}
}
function Add-Map($wt,$pattern,$x,$category,$value=1) {
 $script:rows.Add(@('Wafer Maps',1,1,$wt,$pattern,[double]$x,[double]0,'Monitor Address Count',$category,'598',[double]$value,'Recorded matches'))
}
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Workbook,0,$false)
 $index=$book.Worksheets.Item('Import Index');$loc=$book.Worksheets.Item('Chip Locations')
 $loc.UsedRange.ClearContents()
 for($i=1;$i -le 36;$i++){$loc.Cells.Item($i+1,1).Value2=[double]$i;$loc.Cells.Item($i+1,2).Value2=[double]0}
 for($i=1;$i -le 3;$i++){$loc.Cells.Item($i+1,5).Value2='OK'}
 $index.Range('AA:AD').ClearContents()
 $cats=@(10,11,14,15,16,17)
 for($i=1;$i -le 36;$i++){
   $index.Cells.Item($i+1,27).Value2='WT1'
   $index.Cells.Item($i+1,28).Value2=[double]$i
   $index.Cells.Item($i+1,29).Value2=[double]0
   $index.Cells.Item($i+1,30).Value2=[double]$(if($i -le 6){$cats[$i-1]}else{1})
 }
 $index.Cells.Item(5,10).Value2='FAX7680T0';$index.Cells.Item(5,11).Value2=[double]1
 $index.Range('AT:BA').ClearContents()
 $panels=New-Object 'System.Collections.Generic.List[object]'
 foreach($pattern in @('Scan','March')){
  foreach($phase in 1..3){$panels.Add(@('Wafer Maps','QUAD',('WT'+$phase),$pattern,1,1,1,1))}
 }
 $panel=New-Object 'object[,]' $panels.Count,8
 for($r=0;$r -lt $panels.Count;$r++){for($c=0;$c -lt 8;$c++){$panel[$r,$c]=$panels[$r][$c]}}
 $index.Cells.Item(2,46).Resize($panels.Count,8).Value2=$panel
 $script:rows=New-Object 'System.Collections.Generic.List[object]'
 foreach($x in 7..21){Add-Map 'WT1' 'Scan' $x 'Write Fail' 99}
 foreach($x in 18..32){Add-Map 'WT1' 'Scan' $x 'Read Fail'}
 foreach($x in 20..22){Add-Map 'WT1' 'Scan' $x 'Write&Read Fail'}
 Add-Map 'WT1' 'Scan' 7 'Write Fail' 1000
 Add-Map 'WT1' 'Scan' 33 'Write Fail' 0
 foreach($x in 1..6){Add-Map 'WT1' 'Scan' $x 'Write Fail' 500}
 Add-Map 'WT2' 'Scan' 7 'No monitor item Fail'
 Add-Map 'WT3' 'March' 8 'Read Fail'
 $index.Range('AG:AR').ClearContents()
 $map=New-Object 'object[,]' $rows.Count,12
 for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt 12;$c++){$map[$r,$c]=$rows[$r][$c]}}
 $index.Cells.Item(2,33).Resize($rows.Count,12).Value2=$map
 Run-Report
 $ws=$book.Worksheets.Item('W&R Performance')
 Check ($ws.Range('A1').Value2 -eq "Scan`nW/R Performance") 'Scan merged title'
 Check ($ws.Range('D1').Value2 -eq 'WT1 Chips' -and $ws.Range('H1').Value2 -eq 'WT2 Chips' -and $ws.Range('L1').Value2 -eq 'WT3 Chips') 'WT group headers'
 Check ($ws.Range('A3').Value2 -eq 'FAX768' -and $ws.Range('B3').Value2 -eq '#01' -and $ws.Range('C3').Value2 -eq 30) 'Identity or six-Cat denominator wrong'
 Check ($ws.Range('D3').Value2 -eq .5 -and $ws.Range('E3').Value2 -eq .5 -and $ws.Range('F3').Value2 -eq .1 -and $ws.Range('G3').Value2 -eq 0) 'Deduplication, exclusion or matching counts wrong'
 Check ($ws.Range('K3').Value2 -eq (1.0/30) -and $ws.Range('N8').Value2 -eq 0 -and $ws.Range('M8').Value2 -eq (1.0/30)) 'WT or Pattern mixed'
 Check ($ws.Range('A6').Value2 -eq "March`nW/R Performance") 'March did not output'
 Check ($ws.Range('D3:O3').NumberFormat -eq '0.0%') 'Percentage format'
 Check ($ws.Range('D3').Text -eq '50.0%') 'Extra count text found'
 Check ($ws.UsedRange.Font.Name -eq 'Calibri' -and $ws.UsedRange.Font.Size -eq 11) 'Font'
 Check ($ws.UsedRange.HorizontalAlignment -eq -4108 -and $ws.UsedRange.VerticalAlignment -eq -4108) 'Center alignment'
 $book.SaveCopyAs($Output)
 # Rerunning with one WT must remove WT2/3.
 $loc.Range('E3:E4').ClearContents();$index.Range('AT:BA').ClearContents()
 $index.Range('AT2:BA2').Value2=@('Wafer Maps','QUAD','WT1','Scan',1,1,1,1)
 Run-Report
 Check ($ws.Range('H1').Value2 -eq $null -and $ws.Range('A6').Value2 -ne "March`nW/R Performance") 'Stale WT or Pattern'
 # Zero denominator must display a dash.
 for($i=1;$i -le 36;$i++){$index.Cells.Item($i+1,30).Value2=[double]10}
 Run-Report
 Check ($ws.Range('C3').Value2 -eq 0 -and $ws.Range('D3').Value2 -eq '-') 'Division by zero'
 try{$app.Run("'"+$book.Name+"'!WRPerformance.Invalidate",$book) | Out-Null}catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw}}
 Check ($ws.Range('A1').Value2 -like 'Source updated.*' -and $ws.Range('D3').Value2 -eq $null) 'Stale report invalidation'
 Write-Output 'PASS 3 WT / 2 Pattern percentages, six DC Cats, per-chip/category deduplication, zero counts, overlap, Lot/Wafer, Calibri11 centered, percentage-only cells, rerun3-to1, zero denominator and stale invalidation.'
}finally{if($book){$book.Close($false)};if($app){$app.Quit()}}
