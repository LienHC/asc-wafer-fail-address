param(
 [string]$Workbook=(Join-Path $PSScriptRoot '../templates/SRAM Extra BLK Failure  Analysis_v13.4.xlsm'),
 [string]$Output=(Join-Path ([IO.Path]::GetTempPath()) 'Native_WR_Performance_Test.xlsm')
)
$ErrorActionPreference='Stop';$app=$null;$book=$null
function Check($ok,$message){if(-not $ok){throw $message}}
function Run-Report {
 try{$app.Run("'"+$book.Name+"'!WRTestHarness.BuildReport",$book) | Out-Null}
 catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw}}
}
function Add-Map($wt,$pattern,$x,$category,$value=1) {
 $script:rows.Add(@('Wafer Maps',1,1,$wt,$pattern,[double]$x,[double]0,'Monitor Address Count',$category,'598',[double]$value,'Recorded matches'))
}
function Add-Gray($wt,$pattern,$x) {
 $script:rows.Add(@('Wafer Maps',1,1,$wt,$pattern,[double]$x,[double]0,'Unavailable','','598',$null,'Excluded by Cat'))
}
function Write-Map {
 $index.Range('AG:AR').ClearContents()
 $map=New-Object 'object[,]' $rows.Count,12
 for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt 12;$c++){$map[$r,$c]=$rows[$r][$c]}}
 $index.Cells.Item(2,33).Resize($rows.Count,12).Value2=$map
}
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false
 $book=$app.Workbooks.Open($Workbook,0,$false)
 # Option Private Module cannot be called through Application.Run after a native
 # Excel save. Use a public bridge in this unsaved test instance only.
 $bridge=$book.VBProject.VBComponents.Add(1)
 $bridge.Name='WRTestHarness'
 $bridge.CodeModule.AddFromString(@'
Public Sub BuildReport(ByVal wb As Workbook)
    WRPerformance.BuildReport wb
End Sub
Public Sub Invalidate(ByVal wb As Workbook)
    WRPerformance.Invalidate wb
End Sub
'@)
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
   $index.Cells.Item($i+1,30).Value2=[double]$(if($i -le 6){$cats[$i-1]}elseif($i -eq 36){20}else{1})
 }
 $index.Cells.Item(5,10).Value2='FAX7680T0';$index.Cells.Item(5,11).Value2=[double]1
 $index.Range('AT:BA').ClearContents()
 $panels=New-Object 'System.Collections.Generic.List[object]'
 foreach($pattern in @('Scan','March')){
  foreach($phase in 1..3){if($pattern -eq 'Scan' -and $phase -eq 3){continue};$panels.Add(@('Wafer Maps','QUAD',('WT'+$phase),$pattern,1,1,1,1))}
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
 foreach($x in 1..6){Add-Gray 'WT1' 'Scan' $x}
 Add-Gray 'WT1' 'Scan' 36
 foreach($x in @(1,8,36)){Add-Gray 'WT2' 'Scan' $x}
 Add-Gray 'WT2' 'Scan' 8
 Add-Gray 'WT2' 'March' 2
 Add-Map 'WT2' 'Scan' 7 'No monitor item Fail'
 foreach($mode in @('Write Fail','Read Fail','Write&Read Fail')){
  Add-Map 'WT2' 'Scan' 7 $mode
  Add-Map 'WT2' 'Scan' 8 $mode
 }
 Add-Map 'WT2' 'Scan' 2 'Write Fail'
 Add-Map 'WT2' 'Scan' 9 'Write Fail'
 Add-Map 'WT2' 'Scan' 9 'Read Fail'
 Add-Map 'WT2' 'Scan' 10 'Write Fail'
 Add-Map 'WT2' 'Scan' 10 'Write&Read Fail'
 Add-Map 'WT2' 'Scan' 11 'Read Fail'
 Add-Map 'WT2' 'Scan' 11 'Write&Read Fail'
 Add-Map 'WT3' 'March' 8 'Read Fail'
 # One eligible Chip for each exact combination in this panel.
 $modes=@('Write Fail','Read Fail','Write&Read Fail','No monitor item Fail')
 foreach($mask in 1..15){
  for($bit=0;$bit -lt $modes.Count;$bit++){
   if($mask -band (1 -shl $bit)){Add-Map 'WT2' 'March' ($mask+6) $modes[$bit] (2+$bit)}
  }
 }
 Add-Map 'WT2' 'March' 21 'Write Fail' 999
 Add-Map 'WT2' 'March' 22 'Write Fail' 2
 Add-Map 'WT2' 'March' 22 'Read Fail' 5
 Add-Map 'WT2' 'March' 1 'Read Fail' 500
 Write-Map
 Run-Report
 $ws=$book.Worksheets.Item('W&R Performance')
 Check ($ws.Range('A1').Value2 -eq "Scan`nW/R Performance") 'Scan merged title'
 Check ($ws.Range('D1').Value2 -eq 'WT1 Chips' -and $ws.Range('H1').Value2 -eq 'WT2 Chips' -and $ws.Range('L1').Value2 -eq 'WT3 Chips') 'WT group headers'
 Check ($ws.Range('A3').Value2 -eq 'FAX768' -and $ws.Range('B3').Value2 -eq '#01' -and $ws.Range('C3').Value2 -eq 30) 'Identity or six-Cat denominator wrong'
 Check ($ws.Range('D3').Value2 -eq .5 -and $ws.Range('E3').Value2 -eq .5 -and $ws.Range('F3').Value2 -eq .1 -and $ws.Range('G3').Value2 -eq 0) 'Deduplication, exclusion or matching counts wrong'
 Check ($ws.Range('K3').Value2 -eq (1.0/30) -and $ws.Range('N8').Value2 -eq 0 -and $ws.Range('M8').Value2 -eq (1.0/30)) 'WT or Pattern mixed'
 Check ($ws.Range('A6').Value2 -eq "March`nW/R Performance") 'March did not output'
 Check ($ws.Range('D3:O3').NumberFormat -eq '0.00%') 'Percentage format'
 Check ($ws.Range('D3').Text -eq '50.00%') 'Extra count text found'
 foreach($address in @('A1:O3','A87:D108')){
  $table=$ws.Range($address)
  Check ($table.Font.Name -eq 'Calibri' -and $table.Font.Size -eq 11) "Font for $address"
  Check ($table.HorizontalAlignment -eq -4108 -and $table.VerticalAlignment -eq -4108) "Center alignment for $address"
 }
 Check ($ws.Range('L3').Value2 -eq '-' -and $ws.ChartObjects().Count -eq 5) 'Missing Pattern/WT panel or unexpected chart count'
 Check ($ws.Range('A87').Value2 -eq 'March / WT2 | Fail Count Detail') 'Detailed WT/Pattern title'
 Check ($ws.Range('B88').Value2 -eq 'Fail Count' -and $ws.Range('C88').Value2 -eq 'DC Pass Chips') 'Detail headers'
 Check ($ws.Range('A97').Value2 -eq 'W + R') 'W + R must be distinct from RW'
 Check ($ws.Range('A107').Value2 -eq 'W + R + RW + X') 'Four-mode combination label'
 for($r=89;$r -le 108;$r++){
  $expected=if($r -le 90){9}elseif($r -le 92){8}elseif($r -eq 97){2}elseif($r -eq 108){14}else{1}
  Check ($ws.Cells.Item($r,2).Value2 -eq $expected) "Incorrect numerator on detail row $r"
  Check ($ws.Cells.Item($r,3).Value2 -eq 30) "Incorrect denominator on detail row $r"
  Check ($ws.Cells.Item($r,4).Value2 -eq ($expected / 30.0)) "Incorrect share on detail row $r"
 }
 Check ($ws.Range('D107').Text -eq '3.33%') 'Permutation percentage format'
 Check ($ws.ChartObjects(4).Chart.ChartTitle.Text -like 'March / WT2*30 DC Pass Chips') 'Chart uses the matching panel and denominator'
 Check ($ws.ChartObjects(4).Chart.SeriesCollection(1).Points().Count -eq 19) 'Chart must show four independent metrics and 15 combinations'
 $scan=$book.Worksheets.Item('Scan W-R Performance')
 Check ($scan.Range('D1').Value2 -eq 'WT1 DC Pass Chips' -and $scan.Range('D7').Value2 -eq 'WT2 DC Pass Chips') 'Scan WT blocks'
 Check ($scan.Range('A3').Value2 -eq 'FAX768' -and $scan.Range('B3').Value2 -eq '#01' -and $scan.Range('J2').Value2 -eq 'W+R+W/R') 'Scan table identity'
 Check ($scan.Range('C3').Value2 -eq 29 -and $scan.Range('C9').Value2 -eq 33) 'Scan denominators must use per-WT gray entries and deduplicate XY'
 Check ($scan.Range('A13').Value2 -notlike 'Scan*') 'Missing WT3 Scan unexpectedly included'
 $expected=@(4,3,3,1,1,1,1)
 for($j=0;$j -lt $expected.Count;$j++){
  Check ($scan.Cells.Item(10,4+$j).Value2 -eq $expected[$j]) "Scan numerator at metric $j"
  Check ($scan.Cells.Item(9,4+$j).Value2 -eq ($expected[$j] / 33.0)) "Scan percentage at metric $j"
 }
 Check ($scan.Range('J9').Text -eq '3.03%') 'Scan percentage format'
 $book.SaveCopyAs($Output)
 # Rerunning with one WT must remove WT2/3.
 $loc.Range('E3:E4').ClearContents()
 $index.Cells.Item(3,46).Resize($panels.Count-1,8).ClearContents()
 Run-Report
 Check ($ws.Range('H1').Value2 -eq $null -and $ws.Range('A6').Value2 -ne "March`nW/R Performance") 'Stale WT or Pattern'
 Check ($ws.ChartObjects().Count -eq 1) 'Rerun retained stale charts'
 Check ($scan.Range('D7').Value2 -eq $null) 'Scan rerun retained stale WT'
 $index.Cells.Item(2,49).Value2='March'
 Run-Report
 Check ($scan.Range('A1').Value2 -like '*no Scan Wafer Map*') 'Missing Scan panel'
 $index.Cells.Item(2,49).Value2='Scan'
 # Zero denominator must display a dash.
 for($i=1;$i -le 36;$i++){
  $index.Cells.Item($i+1,30).Value2=[double]10
  Add-Gray 'WT1' 'Scan' $i
 }
 Write-Map
 Run-Report
 Check ($ws.Range('C3').Value2 -eq 0 -and $ws.Range('D3').Value2 -eq '-') 'Division by zero'
 Check ($ws.Range('B17').Value2 -eq 0 -and $ws.Range('C17').Value2 -eq 0 -and $ws.Range('D17').Value2 -eq '-') 'Zero denominator detail'
 Check ($scan.Range('C3').Value2 -eq 0 -and $scan.Range('D3').Value2 -eq '-' -and $scan.Range('D4').Value2 -eq 0) 'Scan zero denominator'
 try{$app.Run("'"+$book.Name+"'!WRTestHarness.Invalidate",$book) | Out-Null}catch [System.ArgumentNullException]{if($_.Exception.ParamName -ne 'ptr'){throw}}
 Check ($ws.Range('A1').Value2 -like 'Source updated.*' -and $ws.Range('D3').Value2 -eq $null) 'Stale report invalidation'
 Check ($ws.ChartObjects().Count -eq 0) 'Invalidation retained stale charts'
 Check ($scan.Range('A1').Value2 -like 'Source updated.*' -and $scan.Range('D4').Value2 -eq $null) 'Stale Scan report invalidation'
 $book.Close($false);$book=$null
 $book=$app.Workbooks.Open($Output,0,$true)
 $ws=$book.Worksheets.Item('W&R Performance')
 Check ($ws.Range('D3').Value2 -eq .5 -and $ws.ChartObjects().Count -eq 5) 'Saved report did not survive reopen'
 Check ($book.Worksheets.Item('Scan W-R Performance').Range('C9').Value2 -eq 33) 'Saved Scan report did not survive reopen'
 Write-Output 'PASS 3 WT / 2 Pattern percentages, six DC Cats, per-chip/category deduplication, zero counts, overlap, all 15 combinations, detail counts and charts, per-WT Scan gray counts and seven metrics, missing panels, rerun, zero denominator and stale invalidation.'
 Write-Output 'PASS saved report reopened in native Excel with counts and charts intact.'
}finally{if($book){$book.Close($false)};if($app){$app.Quit()}}
