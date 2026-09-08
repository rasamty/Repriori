<#
.SYNOPSIS
  Compares a ReportGenerator JsonSummary coverage report against a package's own
  coverage-policy.json, and fails (exit code 1) if either line or branch coverage for
  that specific assembly falls short of policy. See Part 5.2 of
  docs/Repriori_Plan_and_Design.docx for why this exists.

.DESCRIPTION
  Thresholds apply per assembly (one whole package's compiled DLL), never blended
  across the solution and never enforced per individual class -- a package with one
  hard-to-cover class shouldn't have to fake coverage there just to satisfy a rule that
  was never really about that one class. Run from the coverage-gate job in
  .github/workflows/ci.yml, but works identically from a local terminal -- see the
  example below.

.EXAMPLE
  pwsh scripts/Check-CoverageGate.ps1 `
    -SummaryJsonPath TestResults/CoverageReport/Summary.json `
    -PolicyJsonPath src/Repriori.DocProfile/coverage-policy.json
#>
param(
    [Parameter(Mandatory = $true)][string]$SummaryJsonPath,
    [Parameter(Mandatory = $true)][string]$PolicyJsonPath
)

$ErrorActionPreference = "Stop"

$policy = Get-Content $PolicyJsonPath -Raw | ConvertFrom-Json
$summary = Get-Content $SummaryJsonPath -Raw | ConvertFrom-Json

$assembly = $summary.coverage.assemblies | Where-Object { $_.name -eq $policy.assembly }
if (-not $assembly) {
    $names = ($summary.coverage.assemblies | ForEach-Object { $_.name }) -join ", "
    Write-Error "No assembly named '$($policy.assembly)' found in the coverage report. Assemblies present: $names"
    exit 1
}

$lineActual = [math]::Round([double]$assembly.coverage, 1)
$branchActual = [math]::Round([double]$assembly.branchcoverage, 1)
$lineMin = [double]$policy.lineMinPercent
$branchMin = [double]$policy.branchMinPercent

$lineOk = $lineActual -ge $lineMin
$branchOk = $branchActual -ge $branchMin

$report = @"
Coverage gate for $($policy.assembly)
  Line coverage:   $lineActual% (minimum $lineMin%) -- $(if ($lineOk) { "PASS" } else { "FAIL" })
  Branch coverage: $branchActual% (minimum $branchMin%) -- $(if ($branchOk) { "PASS" } else { "FAIL" })
"@

Write-Output $report
if ($env:GITHUB_STEP_SUMMARY) {
    "``````" + "`n$report`n" + "``````" | Add-Content -Path $env:GITHUB_STEP_SUMMARY
}

if (-not ($lineOk -and $branchOk)) {
    $rows = foreach ($class in $assembly.classesinassembly) {
        $classLine = [math]::Round([double]$class.coverage, 1)
        $classBranchRaw = $class.branchcoverage
        $classBranch = if ($null -eq $classBranchRaw) { $null } else { [math]::Round([double]$classBranchRaw, 1) }
        $lineShort = $classLine -lt $lineMin
        $branchShort = ($null -ne $classBranch) -and ($classBranch -lt $branchMin)
        if ($lineShort -or $branchShort) {
            $branchDisplay = if ($null -eq $classBranch) { "n/a" } else { "$classBranch%" }
            "| $($class.name) | $classLine% | $branchDisplay |"
        }
    }

    $failureDetail = @"
## Coverage gate failed: $($policy.assembly)

| Metric | Actual | Required | Result |
|---|---|---|---|
| Line coverage | $lineActual% | $lineMin% | $(if ($lineOk) { "pass" } else { "**FAIL**" }) |
| Branch coverage | $branchActual% | $branchMin% | $(if ($branchOk) { "pass" } else { "**FAIL**" }) |

### Classes below either threshold

| Class | Line % | Branch % |
|---|---|---|
$($rows -join "`n")

Closing this gap is a normal reviewed code change, not something CI writes for you --
see Part 5.2 of docs/Repriori_Plan_and_Design.docx and Part F.5 of
docs/Repriori_Build_Log_and_Test_Plan.docx for the established approach.
"@

    $failureDetail | Out-File -FilePath "coverage-gate-failure.md" -Encoding utf8
    Write-Output $failureDetail
    exit 1
}

exit 0
