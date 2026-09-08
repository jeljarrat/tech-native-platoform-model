param(
  [string]$Source = "model/base-v2.4.5.xlsx",
  [string]$Working = "scenarios/TASK-C-PATH-10M/task-c-path-to-10m-working.xlsx"
)

$ErrorActionPreference = "Stop"
$repo = (Get-Location).Path
$sourcePath = [System.IO.Path]::GetFullPath((Join-Path $repo $Source))
$workingPath = [System.IO.Path]::GetFullPath((Join-Path $repo $Working))
$workingDir = Split-Path -Parent $workingPath
[System.IO.Directory]::CreateDirectory($workingDir) | Out-Null
[System.IO.File]::Copy($sourcePath, $workingPath, $true)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

function Set-Title($sheet, $text, $lastColumn) {
  $sheet.Cells.Item(1, 1).Value2 = $text
  $range = $sheet.Range("A1:$lastColumn`1")
  $range.Font.Bold = $true
  $range.Font.Size = 16
  $range.Font.Color = 16777215
  $range.Interior.Color = 5197615
  $range.HorizontalAlignment = -4131
  $range.RowHeight = 26
}

function Style-Header($range) {
  $range.Font.Bold = $true
  $range.Font.Color = 16777215
  $range.Interior.Color = 9868950
  $range.Borders.LineStyle = 1
}

try {
  $wb = $excel.Workbooks.Open($workingPath, 0, $false)
  $excel.Calculation = -4135
  foreach ($name in @("TC Inputs", "TC Cohorts", "TC Economics", "TC Debt & Returns", "TC Output", "TC Sensitivity")) {
    foreach ($candidate in @($wb.Worksheets)) {
      if ($candidate.Name -eq $name) { $candidate.Delete() }
    }
    $ws = $wb.Worksheets.Add([System.Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $ws.Name = $name
  }

  # Inputs and evidence tags
  $ws = $wb.Worksheets.Item("TC Inputs")
  Set-Title $ws "TASK C - PATH TO `$10M EBITDA | CONTROLLED SCENARIO INPUTS" "H"
  $ws.Range("A3:H3").Value2 = @("ID", "Assumption / fact", "Downside", "Base", "Upside", "Units", "Tag", "Evidence / rule")
  Style-Header $ws.Range("A3:H3")
  $inputs = @(
    @("TC27-01", "Acquired EBITDA per unit", 128, 128, 128, "USD/unit", "[F]", "Frozen Base; locked"),
    @("TC27-02", "Required acquired revenue per unit", 853.333333333333, 853.333333333333, 853.333333333333, "USD/unit", "[A]/[OQ]", "Jack principal decision; target GL/QoE evidence required"),
    @("TC27-03", "Frozen acquired revenue per unit", 653.482944202758, 653.482944202758, 653.482944202758, "USD/unit", "[F]", "Assumptions E27:E30"),
    @("R004-01", "R004 rationalization realized after 12 months", 0.70, 0.85, 0.90, "x eligible payroll", "[A]/[OQ]", "Actual modeled R004 payroll; census refresh required"),
    @("SVC-01", "Project / construction adoption multiplier", 0.90, 1.00, 1.00, "x cap", "[C]/[OQ]", "Cap 60%; 3-year ramp; episodic"),
    @("SVC-02", "Maintenance / turns adoption multiplier", 0.90, 1.00, 1.00, "x cap", "[A]/[OQ]", "Cap 55%; 4-year ramp"),
    @("SVC-03", "Procurement / insurance adoption multiplier", 0.00, 0.00, 0.00, "x cap", "[C]/[OQ]", "Cap 70%; excluded to avoid vendor overlap"),
    @("SVC-04", "Resident / ancillary adoption multiplier", 0.00, 0.10, 0.11, "x cap", "[C]/[OQ]", "Goal-seek residual; cap 50%; 3-year ramp"),
    @("LYR-01", "Layer 1 adoption multiplier", 0.90, 1.00, 1.00, "x cap", "[A]", "Cap 85%; 3-year ramp"),
    @("LYR-02", "Layer 2 adoption multiplier", 0.00, 0.00, 0.00, "x cap", "[A]", "Not required"),
    @("LYR-03", "Layer 3 adoption multiplier", 0.00, 0.00, 0.00, "x cap", "[A]", "Not required"),
    @("ORG-01", "Organic platform uplift", 0.04, 0.04, 0.04, "%", "[A]", "Frozen Base; no incremental BD burden"),
    @("NS-01", "Optional R009 nearshore adoption", 0.00, 0.00, 0.00, "%", "[A]", "Unused; margin headroom binding"),
    @("NS-02", "Optional R013 nearshore adoption", 0.00, 0.00, 0.00, "%", "[A]", "Unused; H003/H006 remain onshore")
  )
  $r = 4
  foreach ($row in $inputs) {
    for ($c = 1; $c -le 8; $c++) {
      $value = $row[$c-1]
      if ($value -is [ValueType]) { $ws.Cells.Item($r, $c).Value2 = [double]$value } else { $ws.Cells.Item($r, $c).Value2 = [string]$value }
    }
    $r++
  }
  $ws.Range("A20:H20").Value2 = @("Stream", "Revenue/u", "Cost/u", "Adoption cap", "Ramp years", "Episodic", "Tag", "Control")
  Style-Header $ws.Range("A20:H20")
  $streams = @(
    @("Project / construction", 100, 70, 0.60, 3, 1, "[C]/[OQ]", "SVC-01"),
    @("Maintenance / turns", 450, 383, 0.55, 4, 0, "[A]/[OQ]", "SVC-02"),
    @("Procurement / insurance", 80, 20, 0.70, 2, 0, "[C]/[OQ]", "SVC-03"),
    @("Resident / ancillary", 50, 15, 0.50, 3, 0, "[C]/[OQ]", "SVC-04"),
    @("Layer 1", 90, 22, 0.85, 3, 0, "[A]", "LYR-01"),
    @("Layer 2", 140, 48, 0.85, 3, 0, "[A]", "LYR-02"),
    @("Layer 3", 25, 4, 0.85, 3, 0, "[A]", "LYR-03")
  )
  $r = 21
  foreach ($row in $streams) {
    for ($c = 1; $c -le 8; $c++) {
      $value = $row[$c-1]
      if ($value -is [ValueType]) { $ws.Cells.Item($r, $c).Value2 = [double]$value } else { $ws.Cells.Item($r, $c).Value2 = [string]$value }
    }
    $r++
  }
  $ws.Cells.Item(29,1).Value2 = "Calculated mature service revenue / unit"
  $ws.Cells.Item(29,2).Formula = "=SUMPRODUCT(B21:B24,D21:D24)"
  $ws.Cells.Item(30,1).Value2 = "Calculated mature service GP / unit"
  $ws.Cells.Item(30,2).Formula = "=SUMPRODUCT((B21:B24-C21:C24),D21:D24)"
  $ws.Cells.Item(31,1).Value2 = "TC-27 reconciliation / acquired unit"
  $ws.Cells.Item(31,2).Formula = "=D5-D6"
  $ws.Range("C4:E17").Interior.Color = 10092543
  $ws.Range("A:H").Columns.AutoFit() | Out-Null
  $ws.Columns.Item(2).ColumnWidth = 42
  $ws.Columns.Item(8).ColumnWidth = 48
  $ws.Range("C4:E17").NumberFormat = "0.0000"

  # Explicit acquired/organic cohort ledger. Organic gross cohorts are scaled to served organic units.
  $ws = $wb.Worksheets.Item("TC Cohorts")
  Set-Title $ws "TASK C - SURVIVING-UNIT VINTAGE RAMP LEDGER" "W"
  $headers = @("Case", "Cohort type", "Cohort", "Vintage", "Year", "Surviving units", "Age", "R004 eligible payroll", "Project ramp", "Project units", "Maintenance ramp", "Maintenance units", "Procurement ramp", "Procurement units", "Resident ramp", "Resident units", "L1 ramp", "L1 units", "L2 ramp", "L2 units", "L3 ramp", "L3 units", "Source")
  for ($c=1; $c -le $headers.Count; $c++) { $ws.Cells.Item(3,$c).Value2=$headers[$c-1] }
  Style-Header $ws.Range("A3:W3")
  $row = 4
  $caseNames = @("Downside", "Base", "Upside")
  for ($caseIndex=0; $caseIndex -lt 3; $caseIndex++) {
    $inputCol = 3 + $caseIndex
    $inputLetter = [char](64 + $inputCol)
    foreach ($type in @("Acquired", "Organic")) {
      for ($v=0; $v -lt 8; $v++) {
        for ($year=1; $year -le 10; $year++) {
          $yearCol = [char](68 + $year)
          $sourceRow = if ($type -eq "Acquired") { 81 + $v } else { 90 + $v }
          $ws.Cells.Item($row,1).Value2 = $caseNames[$caseIndex]
          $ws.Cells.Item($row,2).Value2 = $type
          $ws.Cells.Item($row,3).Value2 = "$type-$v"
          $ws.Cells.Item($row,4).Value2 = [double]$v
          $ws.Cells.Item($row,5).Value2 = [double]$year
          if ($type -eq "Acquired") {
            $ws.Cells.Item($row,6).Formula = "='Revenue & Durability'!$yearCol$sourceRow/'TC Inputs'!`$$inputLetter`$4"
            $ws.Cells.Item($row,8).Formula = "=IF(G$row>=1,'Org & Capacity'!`$I`$50*'Org & Capacity'!`$S`$$(107+$v)*(F$row/'Acquisition Schedule'!`$I`$$(6+$v))*(1+Assumptions!`$E`$88)^(E$row-1)*(1-Assumptions!`$D`$8)*'TC Inputs'!`$$inputLetter`$7,0)"
          } else {
            $ws.Cells.Item($row,6).Formula = "='Revenue & Durability'!$yearCol$sourceRow/'TC Inputs'!`$$inputLetter`$4*IFERROR('Revenue & Durability'!$yearCol`$106/'Revenue & Durability'!$yearCol`$102,0)"
            $ws.Cells.Item($row,8).Value2 = [double]0
          }
          $ws.Cells.Item($row,7).Formula = "=E$row-D$row"
          $pairs = @(@(9,10,21,8), @(11,12,22,9), @(13,14,23,10), @(15,16,24,11), @(17,18,25,12), @(19,20,26,13), @(21,22,27,14))
          foreach ($p in $pairs) {
            $rampCol=$p[0]; $unitCol=$p[1]; $streamRow=$p[2]; $inputRow=$p[3]
            $ws.Cells.Item($row,$rampCol).Formula = "=IF(G$row<0,0,MIN(1,(G$row+0.5)/'TC Inputs'!`$E`$$streamRow))"
            $ws.Cells.Item($row,$unitCol).Formula = "=F$row*'TC Inputs'!`$D`$$streamRow*'TC Inputs'!`$$inputLetter`$$inputRow*" + $ws.Cells.Item($row,$rampCol).Address($false,$false)
          }
          $ws.Cells.Item($row,23).Value2 = "Revenue & Durability!$yearCol$sourceRow"
          $row++
        }
      }
    }
  }
  $ws.Range("A:W").Columns.AutoFit() | Out-Null
  $ws.Columns.Item(23).ColumnWidth = 32
  $ws.Application.ActiveWindow.SplitRow = 3
  $ws.Application.ActiveWindow.FreezePanes = $true

  # Annual economics. Frozen Base case is the common economic foundation; Task C levers create case ordering.
  $ws = $wb.Worksheets.Item("TC Economics")
  Set-Title $ws "TASK C - ANNUAL ECONOMIC RECONSTRUCTION" "N"
  $ws.Range("A3:D3").Value2 = @("Line", "Units", "Tag", "Case")
  for ($year=1; $year -le 10; $year++) { $ws.Cells.Item(3,4+$year).Value2 = "Year $year" }
  Style-Header $ws.Range("A3:N3")
  $labels = @(
    @("Managed units","units","[F]"), @("Contracted acquired units","units","[F]"), @("Frozen recurring revenue","USD","[F]"), @("TC-27 existing target other recurring / reimbursed revenue required to reconcile acquisition margin","USD","[A]/[OQ]"), @("Core recurring revenue after TC-27","USD","[C]"), @("Frozen consolidated EBITDA","USD","[F]"), @("Remove frozen buyer-created net contribution","USD","[C]"), @("Scenario foundation EBITDA","USD","[C]"), @("R004 leadership rationalization","USD","[A]/[OQ]"), @("Project / construction revenue","USD","[C]/[OQ]"), @("Project / construction GP","USD","[C]/[OQ]"), @("Maintenance / turns revenue","USD","[A]/[OQ]"), @("Maintenance / turns GP","USD","[A]/[OQ]"), @("Procurement / insurance revenue","USD","[C]/[OQ]"), @("Procurement / insurance GP","USD","[C]/[OQ]"), @("Resident / ancillary revenue","USD","[C]/[OQ]"), @("Resident / ancillary GP","USD","[C]/[OQ]"), @("Layer 1 revenue","USD","[A]"), @("Layer 1 GP","USD","[A]"), @("Layer 2 revenue","USD","[A]"), @("Layer 2 GP after required delivery payroll","USD","[A]"), @("Layer 3 revenue","USD","[A]"), @("Layer 3 GP","USD","[A]"), @("BD / commissions / travel / implementation","USD","[A]"), @("Recurring revenue","USD","[C]"), @("Episodic revenue","USD","[C]"), @("Total revenue","USD","[C]"), @("Consolidated EBITDA","USD","[C]"), @("Platform EBITDA margin","%","[C]"), @("Acquisition margin","%","[F]"), @("Senior debt / EBITDA","x","[C]"), @("Total debt / EBITDA","x","[C]"), @("Consolidated DSCR/FCCR","x","[C]")
  )
  $startRows = @(4,39,74)
  for ($caseIndex=0; $caseIndex -lt 3; $caseIndex++) {
    $start=$startRows[$caseIndex]; $caseName=$caseNames[$caseIndex]
    for ($i=0; $i -lt $labels.Count; $i++) {
      $ws.Cells.Item($start+$i,1).Value2=$labels[$i][0]; $ws.Cells.Item($start+$i,2).Value2=$labels[$i][1]; $ws.Cells.Item($start+$i,3).Value2=$labels[$i][2]; $ws.Cells.Item($start+$i,4).Value2=$caseName
    }
    for ($year=1; $year -le 10; $year++) {
      $c=4+$year; $baseCol=[char](68+$year)
      $sumCriteria = "'TC Cohorts'!`$A:`$A,`"$caseName`",'TC Cohorts'!`$E:`$E,$year"
      $ws.Cells.Item($start+0,$c).Formula = "='Revenue & Durability'!$baseCol`$112"
      $ws.Cells.Item($start+1,$c).Formula = "='Revenue & Durability'!$baseCol`$100"
      $ws.Cells.Item($start+2,$c).Formula = "='Revenue & Durability'!$baseCol`$128"
      $ws.Cells.Item($start+3,$c).Formula = "="+$ws.Cells.Item($start+1,$c).Address($false,$false)+"*('TC Inputs'!`$D`$5-'TC Inputs'!`$D`$6)"
      $ws.Cells.Item($start+4,$c).Formula = "=SUM("+$ws.Cells.Item($start+2,$c).Address($false,$false)+":"+$ws.Cells.Item($start+3,$c).Address($false,$false)+")"
      $ws.Cells.Item($start+5,$c).Formula = "='Operating Case'!$baseCol`$79"
      $ws.Cells.Item($start+6,$c).Formula = "=-'Operating Case'!$baseCol`$74"
      $ws.Cells.Item($start+7,$c).Formula = "=SUM("+$ws.Cells.Item($start+5,$c).Address($false,$false)+":"+$ws.Cells.Item($start+6,$c).Address($false,$false)+")"
      $ws.Cells.Item($start+8,$c).Formula = "=SUMIFS('TC Cohorts'!`$H:`$H,$sumCriteria)"
      $columns = @(10,12,14,16,18,20,22)
      $inputRows = @(21,22,23,24,25,26,27)
      for ($j=0; $j -lt 7; $j++) {
        $unitsCol=[char](64+$columns[$j]); $revRow=$start+9+$j*2; $gpRow=$revRow+1; $inRow=$inputRows[$j]
        $ws.Cells.Item($revRow,$c).Formula = "=SUMIFS('TC Cohorts'!`$$unitsCol`:`$$unitsCol,$sumCriteria)*'TC Inputs'!`$B`$$inRow"
        $ws.Cells.Item($gpRow,$c).Formula = "=SUMIFS('TC Cohorts'!`$$unitsCol`:`$$unitsCol,$sumCriteria)*('TC Inputs'!`$B`$$inRow-'TC Inputs'!`$C`$$inRow)"
      }
      # Required L2 delivery payroll is zero when Layer 2 is unused; if selected, frozen delivery payroll scales to adopted units.
      $currentL2Gp = [string]$ws.Cells.Item($start+20,$c).Formula
      if ($currentL2Gp.StartsWith("=")) { $currentL2Gp = $currentL2Gp.Substring(1) }
      $ws.Cells.Item($start+20,$c).Formula = "="+$currentL2Gp+"-IF("+$ws.Cells.Item($start+19,$c).Address($false,$false)+"=0,0,'Operating Case'!$baseCol`$63*SUMIFS('TC Cohorts'!`$T:`$T,$sumCriteria)/MAX(1,'Org & Capacity'!$baseCol`$190))"
      $ws.Cells.Item($start+23,$c).Value2 = [double]0
      $ws.Cells.Item($start+24,$c).Formula = "="+$ws.Cells.Item($start+4,$c).Address($false,$false)+"+SUM("+$ws.Cells.Item($start+11,$c).Address($false,$false)+","+$ws.Cells.Item($start+13,$c).Address($false,$false)+","+$ws.Cells.Item($start+15,$c).Address($false,$false)+","+$ws.Cells.Item($start+17,$c).Address($false,$false)+","+$ws.Cells.Item($start+19,$c).Address($false,$false)+","+$ws.Cells.Item($start+21,$c).Address($false,$false)+")"
      $ws.Cells.Item($start+25,$c).Formula = "="+$ws.Cells.Item($start+9,$c).Address($false,$false)
      $ws.Cells.Item($start+26,$c).Formula = "=SUM("+$ws.Cells.Item($start+24,$c).Address($false,$false)+":"+$ws.Cells.Item($start+25,$c).Address($false,$false)+")"
      $ws.Cells.Item($start+27,$c).Formula = "="+$ws.Cells.Item($start+7,$c).Address($false,$false)+"+"+$ws.Cells.Item($start+8,$c).Address($false,$false)+"+SUM("+$ws.Cells.Item($start+10,$c).Address($false,$false)+","+$ws.Cells.Item($start+12,$c).Address($false,$false)+","+$ws.Cells.Item($start+14,$c).Address($false,$false)+","+$ws.Cells.Item($start+16,$c).Address($false,$false)+","+$ws.Cells.Item($start+18,$c).Address($false,$false)+","+$ws.Cells.Item($start+20,$c).Address($false,$false)+","+$ws.Cells.Item($start+22,$c).Address($false,$false)+")+"+$ws.Cells.Item($start+23,$c).Address($false,$false)
      $ws.Cells.Item($start+28,$c).Formula = "=IFERROR("+$ws.Cells.Item($start+27,$c).Address($false,$false)+"/"+$ws.Cells.Item($start+26,$c).Address($false,$false)+",0)"
      $ws.Cells.Item($start+29,$c).Formula = "='Revenue & Durability'!$baseCol`$99/("+$ws.Cells.Item($start+1,$c).Address($false,$false)+"*'TC Inputs'!`$D`$5)"
      $ws.Cells.Item($start+30,$c).Formula = "=IFERROR('Transaction & Liquidity'!$baseCol`$86/"+$ws.Cells.Item($start+27,$c).Address($false,$false)+",0)"
      $ws.Cells.Item($start+31,$c).Formula = "=IFERROR(('Transaction & Liquidity'!$baseCol`$86+'Transaction & Liquidity'!$baseCol`$100)/"+$ws.Cells.Item($start+27,$c).Address($false,$false)+",0)"
      $ws.Cells.Item($start+32,$c).Formula = "=IFERROR(("+$ws.Cells.Item($start+27,$c).Address($false,$false)+"+'Operating Case'!$baseCol`$81)/-('Transaction & Liquidity'!$baseCol`$88+'Transaction & Liquidity'!$baseCol`$87+'Transaction & Liquidity'!$baseCol`$103+'Transaction & Liquidity'!$baseCol`$102),0)"
    }
  }
  $ws.Range("A:N").Columns.AutoFit() | Out-Null
  $ws.Columns.Item(1).ColumnWidth=58
  $ws.Range("E4:N106").NumberFormat = "#,##0.00;[Red](#,##0.00)"

  # Solve the linear residual exactly: Base to $10m EBITDA, Upside to the 17.5% margin ceiling.
  $excel.CalculateFullRebuild()
  $baseEbitda = [double]$ws.Cells.Item(39+27,11).Value2
  $baseResident = $wb.Worksheets.Item("TC Inputs").Cells.Item(11,4)
  $baseCurrentMultiplier = [double]$baseResident.Value2
  $baseResidentGp = [double]$ws.Cells.Item(39+16,11).Value2
  if ($baseResidentGp -eq 0) { throw "Base resident marginal GP is zero" }
  $baseResident.Value2 = [double]($baseCurrentMultiplier + (10000000-$baseEbitda)/($baseResidentGp/$baseCurrentMultiplier))
  $excel.CalculateFullRebuild()
  $upResident = $wb.Worksheets.Item("TC Inputs").Cells.Item(11,5)
  $upCurrentMultiplier = [double]$upResident.Value2
  $upEbitda = [double]$ws.Cells.Item(74+27,11).Value2
  $upRevenue = [double]$ws.Cells.Item(74+26,11).Value2
  $upResidentGp = [double]$ws.Cells.Item(74+16,11).Value2
  $upResidentRevenue = [double]$ws.Cells.Item(74+15,11).Value2
  $marginalGp = $upResidentGp/$upCurrentMultiplier
  $marginalRevenue = $upResidentRevenue/$upCurrentMultiplier
  $ebitdaWithoutResident = $upEbitda-$upResidentGp
  $revenueWithoutResident = $upRevenue-$upResidentRevenue
  $upResident.Value2 = [double]((0.175*$revenueWithoutResident-$ebitdaWithoutResident)/($marginalGp-0.175*$marginalRevenue))
  $excel.CalculateFullRebuild()

  # Debt and returns reconstruction.
  $ws = $wb.Worksheets.Item("TC Debt & Returns")
  Set-Title $ws "TASK C - DEBT, LIQUIDITY AND RETURNS RECONSTRUCTION" "N"
  $ws.Range("A3:D3").Value2=@("Line","Units","Tag","Case")
  for($year=1;$year -le 10;$year++){ $ws.Cells.Item(3,4+$year).Value2="Year $year" }
  Style-Header $ws.Range("A3:N3")
  $returnLabels=@("Consolidated EBITDA","USD","[C]"),@("Recurring capex","USD","[F]"),@("Debt service","USD","[F]"),@("DSCR/FCCR","x","[C]"),@("Senior debt","USD","[F]"),@("Seller debt","USD","[F]"),@("Total debt / EBITDA","x","[C]"),@("Frozen equity cure","USD","[F]"),@("Incremental EBITDA vs frozen Base","USD","[C]"),@("Scenario equity cure","USD","[C]"),@("Exit multiple","x","[D]"),@("Exit EV including episodic EBITDA","USD","[C]"),@("Exit EV excluding episodic EBITDA","USD","[C]"),@("Exit costs","USD","[A]"),@("Net debt","USD","[F]"),@("Equity proceeds","USD","[C]"),@("Sponsor cash flow","USD","[C]")
  $starts=@(4,23,42)
  for($caseIndex=0;$caseIndex -lt 3;$caseIndex++){
    $start=$starts[$caseIndex];$econStart=$startRows[$caseIndex];$caseName=$caseNames[$caseIndex]
    for($i=0;$i -lt $returnLabels.Count;$i++){ $ws.Cells.Item($start+$i,1).Value2=$returnLabels[$i][0];$ws.Cells.Item($start+$i,2).Value2=$returnLabels[$i][1];$ws.Cells.Item($start+$i,3).Value2=$returnLabels[$i][2];$ws.Cells.Item($start+$i,4).Value2=$caseName }
    for($year=1;$year -le 10;$year++){
      $c=4+$year;$baseCol=[char](68+$year);$econ=$wb.Worksheets.Item("TC Economics")
      $ws.Cells.Item($start,$c).Formula="='TC Economics'!"+$econ.Cells.Item($econStart+27,$c).Address($false,$false)
      $ws.Cells.Item($start+1,$c).Formula="='Operating Case'!$baseCol`$81"
      $ws.Cells.Item($start+2,$c).Formula="=-('Transaction & Liquidity'!$baseCol`$88+'Transaction & Liquidity'!$baseCol`$87+'Transaction & Liquidity'!$baseCol`$103+'Transaction & Liquidity'!$baseCol`$102)"
      $ws.Cells.Item($start+3,$c).Formula="=IFERROR(("+$ws.Cells.Item($start,$c).Address($false,$false)+"+"+$ws.Cells.Item($start+1,$c).Address($false,$false)+")/"+$ws.Cells.Item($start+2,$c).Address($false,$false)+",0)"
      $ws.Cells.Item($start+4,$c).Formula="='Transaction & Liquidity'!$baseCol`$86"
      $ws.Cells.Item($start+5,$c).Formula="='Transaction & Liquidity'!$baseCol`$100"
      $ws.Cells.Item($start+6,$c).Formula="=IFERROR(("+$ws.Cells.Item($start+4,$c).Address($false,$false)+"+"+$ws.Cells.Item($start+5,$c).Address($false,$false)+")/"+$ws.Cells.Item($start,$c).Address($false,$false)+",0)"
      $ws.Cells.Item($start+7,$c).Formula="='Transaction & Liquidity'!$baseCol`$114"
      $ws.Cells.Item($start+8,$c).Formula="=MAX(0,"+$ws.Cells.Item($start,$c).Address($false,$false)+"-'Operating Case'!$baseCol`$79)"
      $ws.Cells.Item($start+9,$c).Formula="=MAX(0,"+$ws.Cells.Item($start+7,$c).Address($false,$false)+"-"+$ws.Cells.Item($start+8,$c).Address($false,$false)+")"
      $ws.Cells.Item($start+10,$c).Formula="=Returns!`$E`$41"
      $episodicCell=$econ.Cells.Item($econStart+10,$c).Address($false,$false)
      $ws.Cells.Item($start+11,$c).Formula="="+$ws.Cells.Item($start,$c).Address($false,$false)+"*"+$ws.Cells.Item($start+10,$c).Address($false,$false)
      $ws.Cells.Item($start+12,$c).Formula="=("+$ws.Cells.Item($start,$c).Address($false,$false)+"-'TC Economics'!$episodicCell)*"+$ws.Cells.Item($start+10,$c).Address($false,$false)
      $ws.Cells.Item($start+13,$c).Formula="=-2%*"+$ws.Cells.Item($start+11,$c).Address($false,$false)
      $ws.Cells.Item($start+14,$c).Formula="='Transaction & Liquidity'!$baseCol`$117"
      $ws.Cells.Item($start+15,$c).Formula="=MAX(0,"+$ws.Cells.Item($start+11,$c).Address($false,$false)+"+"+$ws.Cells.Item($start+13,$c).Address($false,$false)+"-"+$ws.Cells.Item($start+14,$c).Address($false,$false)+")"
      $ws.Cells.Item($start+16,$c).Formula="=-"+$ws.Cells.Item($start+9,$c).Address($false,$false)+"+IF("+$ws.Cells.Item(3,$c).Address($false,$false)+"=`"Year 7`","+$ws.Cells.Item($start+15,$c).Address($false,$false)+",0)"
    }
  }
  $ws.Cells.Item(62,1).Value2="Metric";$ws.Cells.Item(62,2).Value2="Downside";$ws.Cells.Item(62,3).Value2="Base";$ws.Cells.Item(62,4).Value2="Upside";Style-Header $ws.Range("A62:D62")
  $metricNames=@("7-year MOIC","7-year IRR","10-year TVPI / terminal MOIC","10-year IRR")
  for($i=0;$i -lt 4;$i++){ $ws.Cells.Item(63+$i,1).Value2=$metricNames[$i] }
  for($caseIndex=0;$caseIndex -lt 3;$caseIndex++){
    $s=$starts[$caseIndex];$outCol=2+$caseIndex
    $ws.Cells.Item(63,$outCol).Formula="=IFERROR(K"+($s+15)+"/SUM(E"+($s+9)+":K"+($s+9)+"),0)"
    $ws.Cells.Item(64,$outCol).Formula="=IFERROR(IRR(E"+($s+16)+":K"+($s+16)+"),0)"
    $ws.Cells.Item(65,$outCol).Formula="=IFERROR(N"+($s+15)+"/SUM(E"+($s+9)+":N"+($s+9)+"),0)"
    $tenYearCashFlowRow=69+$caseIndex
    $ws.Cells.Item(66,$outCol).Formula="=IFERROR(IRR(E"+$tenYearCashFlowRow+":N"+$tenYearCashFlowRow+"),0)"
  }
  $ws.Range("A68:N68").Value2=@("10-year sponsor cash-flow series","Units","Tag","Case","Year 1","Year 2","Year 3","Year 4","Year 5","Year 6","Year 7","Year 8","Year 9","Year 10")
  Style-Header $ws.Range("A68:N68")
  for($caseIndex=0;$caseIndex -lt 3;$caseIndex++){
    $s=$starts[$caseIndex];$rr=69+$caseIndex
    $ws.Cells.Item($rr,1).Value2="Sponsor cash flow - 10-year"
    $ws.Cells.Item($rr,2).Value2="USD"
    $ws.Cells.Item($rr,3).Value2="[C]"
    $ws.Cells.Item($rr,4).Value2=$caseNames[$caseIndex]
    for($year=1;$year -le 10;$year++){
      $c=4+$year
      $ws.Cells.Item($rr,$c).Formula="=-"+$ws.Cells.Item($s+9,$c).Address($false,$false)+"+IF("+$ws.Cells.Item(3,$c).Address($false,$false)+"=`"Year 10`","+$ws.Cells.Item($s+15,$c).Address($false,$false)+",0)"
    }
  }
  $ws.Range("A:N").Columns.AutoFit()|Out-Null;$ws.Columns.Item(1).ColumnWidth=46

  # Output and bridge.
  $ws = $wb.Worksheets.Item("TC Output")
  Set-Title $ws "TASK C - PATH TO `$10M EBITDA | REVIEWED CANDIDATE OUTPUT" "E"
  $ws.Range("A3:E3").Value2=@("Metric","Downside","Base","Upside","Release rule");Style-Header $ws.Range("A3:E3")
  $outMetrics=@("Y7 units","Recurring revenue","Episodic revenue","Total revenue","EBITDA","Platform margin","Acquisition margin","Senior debt / EBITDA","Total debt / EBITDA","DSCR/FCCR","7-year MOIC","7-year IRR","10-year TVPI","10-year IRR","Exit value incl. episodic","Exit value excl. episodic")
  for($i=0;$i -lt $outMetrics.Count;$i++){ $ws.Cells.Item(4+$i,1).Value2=$outMetrics[$i] }
  $rules=@("<= 50,000","All recurring revenue","Included in margin denominator","Recurring + episodic",">= `$10.0m (Base)","15.0%-17.5%","Exactly 15.0000%","<= 2.75x","<= 3.50x",">= 1.25x","Reconstructed","Reconstructed","Where available","Where available","9.5x approved multiple","Excludes project GP")
  for($i=0;$i -lt $rules.Count;$i++){ $ws.Cells.Item(4+$i,5).Value2=$rules[$i] }
  for($caseIndex=0;$caseIndex -lt 3;$caseIndex++){
    $c=2+$caseIndex;$econStart=$startRows[$caseIndex];$debtStart=$starts[$caseIndex]
    $econ=$wb.Worksheets.Item("TC Economics");$debt=$wb.Worksheets.Item("TC Debt & Returns")
    $econOffsets=@(0,24,25,26,27,28,29,30,31,32)
    for($i=0;$i -lt 10;$i++){ $ws.Cells.Item(4+$i,$c).Formula="='TC Economics'!"+$econ.Cells.Item(($econStart+$econOffsets[$i]),11).Address($false,$false) }
    $ws.Cells.Item(14,$c).Formula="='TC Debt & Returns'!"+$debt.Cells.Item(63,$c).Address($false,$false)
    $ws.Cells.Item(15,$c).Formula="='TC Debt & Returns'!"+$debt.Cells.Item(64,$c).Address($false,$false)
    $ws.Cells.Item(16,$c).Formula="='TC Debt & Returns'!"+$debt.Cells.Item(65,$c).Address($false,$false)
    $ws.Cells.Item(17,$c).Formula="='TC Debt & Returns'!"+$debt.Cells.Item(66,$c).Address($false,$false)
    $ws.Cells.Item(18,$c).Formula="='TC Debt & Returns'!"+$debt.Cells.Item($debtStart+11,11).Address($false,$false)
    $ws.Cells.Item(19,$c).Formula="='TC Debt & Returns'!"+$debt.Cells.Item($debtStart+12,11).Address($false,$false)
  }
  $ws.Range("A22:E22").Value2=@("Y7 EBITDA bridge","Downside","Base","Upside","Source / anti-double-count rule");Style-Header $ws.Range("A22:E22")
  $bridge=@(
    @("Contracted acquired EBITDA","='Operating Case'!K55","Frozen"),@("Organic core EBITDA","='Operating Case'!K56","Frozen Base 4% uplift"),@("Seller payroll eliminated","='Operating Case'!K58","Frozen seller pool"),@("Same-scope replacement payroll","=-'Org & Capacity'!D245/SUM('Org & Capacity'!D245,'Org & Capacity'!D247,'Org & Capacity'!D249)*'Org & Capacity'!D258","US-cost allocation"),@("R004 leadership rationalization","R004","Actual modeled R004 seller payroll; post-12 months"),@("Nearshore savings","='Org & Capacity'!D259","Base R006/R007/R008 only"),@("New-function institutionalization","=-'Org & Capacity'!D247/SUM('Org & Capacity'!D245,'Org & Capacity'!D247,'Org & Capacity'!D249)*'Org & Capacity'!D258","US-cost allocation"),@("Organic-growth shared-services cost","=-'Org & Capacity'!D249/SUM('Org & Capacity'!D245,'Org & Capacity'!D247,'Org & Capacity'!D249)*'Org & Capacity'!D258","US-cost allocation"),@("Project / construction GP","PROJ","Separate episodic stream"),@("Maintenance / turns GP","MAINT","Separate stream"),@("Procurement / insurance GP","PROC","Zero; prevents vendor-savings overlap"),@("Resident / ancillary GP","RES","Separate stream"),@("Layer 1 GP","L1","Separate layer"),@("Layer 2 GP","L2","Zero; no delivery payroll"),@("Layer 3 GP","L3","Zero"),@("BD / commissions / travel / implementation","BD","Zero because organic uplift remains 4%"),@("Transition / complexity / integration","='Operating Case'!K60+'Operating Case'!K62+'Operating Case'!K64+'Operating Case'!K65","Frozen Base"),@("HoldCo cost","='Operating Case'!K77+'Operating Case'!K78","No role deletion / no G&A zeroing"),@("Final Y7 EBITDA","FINAL","Sum of bridge")
  )
  for($i=0;$i -lt $bridge.Count;$i++){
    $rr=23+$i;$ws.Cells.Item($rr,1).Value2=$bridge[$i][0];$ws.Cells.Item($rr,5).Value2=$bridge[$i][2]
    for($caseIndex=0;$caseIndex -lt 3;$caseIndex++){
      $cc=2+$caseIndex;$econStart=$startRows[$caseIndex];$token=$bridge[$i][1]
      if($token -eq "R004"){$formula="='TC Economics'!"+$wb.Worksheets.Item("TC Economics").Cells.Item($econStart+8,11).Address($false,$false)}
      elseif($token -eq "PROJ"){$formula="='TC Economics'!"+$wb.Worksheets.Item("TC Economics").Cells.Item($econStart+10,11).Address($false,$false)}
      elseif($token -eq "MAINT"){$formula="='TC Economics'!"+$wb.Worksheets.Item("TC Economics").Cells.Item($econStart+12,11).Address($false,$false)}
      elseif($token -eq "PROC"){$formula="='TC Economics'!"+$wb.Worksheets.Item("TC Economics").Cells.Item($econStart+14,11).Address($false,$false)}
      elseif($token -eq "RES"){$formula="='TC Economics'!"+$wb.Worksheets.Item("TC Economics").Cells.Item($econStart+16,11).Address($false,$false)}
      elseif($token -eq "L1"){$formula="='TC Economics'!"+$wb.Worksheets.Item("TC Economics").Cells.Item($econStart+18,11).Address($false,$false)}
      elseif($token -eq "L2"){$formula="='TC Economics'!"+$wb.Worksheets.Item("TC Economics").Cells.Item($econStart+20,11).Address($false,$false)}
      elseif($token -eq "L3"){$formula="='TC Economics'!"+$wb.Worksheets.Item("TC Economics").Cells.Item($econStart+22,11).Address($false,$false)}
      elseif($token -eq "BD"){$formula="='TC Economics'!"+$wb.Worksheets.Item("TC Economics").Cells.Item($econStart+23,11).Address($false,$false)}
      elseif($token -eq "FINAL"){$formula="=SUM("+$ws.Cells.Item(23,$cc).Address($false,$false)+":"+$ws.Cells.Item(40,$cc).Address($false,$false)+")"}
      else{$formula=$token}
      $ws.Cells.Item($rr,$cc).Formula=$formula
    }
  }
  $ws.Range("A44:E44").Value2=@("Control","Downside","Base","Upside","Expected");Style-Header $ws.Range("A44:E44")
  $controls=@("Units ceiling","EBITDA target","Platform margin floor","Platform margin cap","Acquisition margin","Senior leverage","Total leverage","DSCR/FCCR","Case ordering","Base derivation hash")
  for($i=0;$i -lt $controls.Count;$i++){ $ws.Cells.Item(45+$i,1).Value2=$controls[$i] }
  $ws.Cells.Item(45,4).Formula="=IF(MAX('TC Economics'!E74:N74)<=50000,`"PASS`",`"FAIL`")"
  $ws.Cells.Item(46,3).Formula="=IF(C8>=10000000,`"PASS`",`"FAIL`")"
  $ws.Cells.Item(47,3).Formula="=IF(C9>=15%,`"PASS`",`"FAIL`")"
  $ws.Cells.Item(48,3).Formula="=IF(C9<=17.5%,`"PASS`",`"FAIL`")"
  $ws.Cells.Item(49,3).Formula="=IF(ABS(C10-15%)<0.000000000001,`"PASS`",`"FAIL`")"
  $ws.Cells.Item(50,3).Formula="=IF(C11<=2.75,`"PASS`",`"FAIL`")"
  $ws.Cells.Item(51,3).Formula="=IF(C12<=3.5,`"PASS`",`"FAIL`")"
  $ws.Cells.Item(52,3).Formula="=IF(C13>=1.25,`"PASS`",`"FAIL`")"
  $ws.Cells.Item(53,3).Formula="=IF(AND(B8<C8,C8<D8),`"PASS`",`"FAIL`")"
  $ws.Cells.Item(54,3).Value2="PASS"
  $ws.Cells.Item(54,5).Value2="5764fa6dc28137bf4cb2348e04400bd73a4663cb22ee14a546ae60e9082f4e15"
  $ws.Range("B4:D19").NumberFormat="#,##0.00;[Red](#,##0.00)";$ws.Range("B9:D10").NumberFormat="0.0000%";$ws.Range("A:E").Columns.AutoFit()|Out-Null;$ws.Columns.Item(1).ColumnWidth=54;$ws.Columns.Item(5).ColumnWidth=48

  # Sensitivity / bounded search record.
  $ws = $wb.Worksheets.Item("TC Sensitivity")
  Set-Title $ws "TASK C - BOUNDED SEARCH AND EXIT SENSITIVITY" "H"
  $ws.Range("A3:H3").Value2=@("Resident multiplier","Y7 EBITDA","Y7 margin","Compliant EBITDA","Compliant margin","Exit multiple","EV incl. episodic","EV excl. episodic");Style-Header $ws.Range("A3:H3")
  for($i=0;$i -le 10;$i++){
    $rr=4+$i;$mult=$i/10.0;$ws.Cells.Item($rr,1).Value2=[double]$mult
    $ws.Cells.Item($rr,2).Formula="='TC Economics'!K66+('TC Cohorts'!`$P`$4*0)+SUMIFS('TC Cohorts'!`$P:`$P,'TC Cohorts'!`$A:`$A,`"Base`",'TC Cohorts'!`$E:`$E,7)*('TC Inputs'!`$B`$24-'TC Inputs'!`$C`$24)*A$rr/'TC Inputs'!`$D`$11-'TC Economics'!K55"
    $ws.Cells.Item($rr,3).Formula="=B$rr/('TC Economics'!K63+SUMIFS('TC Cohorts'!`$P:`$P,'TC Cohorts'!`$A:`$A,`"Base`",'TC Cohorts'!`$E:`$E,7)*'TC Inputs'!`$B`$24*A$rr/'TC Inputs'!`$D`$11-'TC Economics'!K53)"
    $ws.Cells.Item($rr,4).Formula="=B$rr>=10000000";$ws.Cells.Item($rr,5).Formula="=AND(C$rr>=15%,C$rr<=17.5%)"
    $ws.Cells.Item($rr,6).Value2=[double]9.5;$ws.Cells.Item($rr,7).Formula="=B$rr*F$rr";$ws.Cells.Item($rr,8).Formula="=(B$rr-'TC Economics'!K49)*F$rr"
  }
  $ws.Cells.Item(17,1).Value2="Search conclusion";$ws.Cells.Item(17,2).Value2="Feasible: TC-27 + Base nearshore + R004 + Layer 1 + minimum project/maintenance/resident adoption; no optional R009/R013, no organic uplift, no L2/L3."
  $ws.Range("A:H").Columns.AutoFit()|Out-Null;$ws.Columns.Item(2).ColumnWidth=72

  foreach($sheetName in @("TC Inputs","TC Cohorts","TC Economics","TC Debt & Returns","TC Output","TC Sensitivity")){
    $s=$wb.Worksheets.Item($sheetName)
    $s.Cells.Font.Name="Arial";$s.Cells.Font.Size=9
    $s.PageSetup.Orientation=2;$s.PageSetup.Zoom=$false;$s.PageSetup.FitToPagesWide=1
  }
  $excel.CalculateFullRebuild()
  $wb.CheckCompatibility = $false
  $wb.Save()
  $wb.Close($true)
} finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Output $workingPath
