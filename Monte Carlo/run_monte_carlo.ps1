param(
  [string]$Config = "mc_assumptions.json",
  [string]$Input = "",
  [string]$OutputDir = "",
  [int]$Simulations = 0,
  [int]$Seed = 0
)

$scriptPath = Join-Path $PSScriptRoot "mc_pipeline.py"
$configPath = Join-Path $PSScriptRoot $Config

$argsList = @($scriptPath, "--config", $configPath)
if ($Input -ne "") { $argsList += @("--input", $Input) }
if ($OutputDir -ne "") { $argsList += @("--output-dir", $OutputDir) }
if ($Simulations -gt 0) { $argsList += @("--n-sim", $Simulations) }
if ($Seed -gt 0) { $argsList += @("--seed", $Seed) }

python @argsList
