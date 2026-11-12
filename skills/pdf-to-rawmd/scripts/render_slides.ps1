param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [Parameter(Mandatory = $true)][string]$Slides,
    [ValidateRange(320, 7680)][int]$Width = 1600
)

$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $InputPath).Path
$destination = [System.IO.Path]::GetFullPath($OutputDir)
if ([System.IO.Path]::GetExtension($source) -ne '.pptx') { throw 'Expected .pptx input.' }
if (Test-Path -LiteralPath $destination) { throw 'OutputDir must be a new directory.' }
if ($Slides -notmatch '^\s*[1-9][0-9]*(?:\s*,\s*[1-9][0-9]*)*\s*$') {
    throw 'Use comma-separated slide numbers, e.g. 3,6,14.'
}
$numbers = @($Slides.Split(',') | ForEach-Object { [int]$_.Trim() } | Sort-Object -Unique)
if ($numbers.Count -gt 15) { throw 'Select at most 15 slides.' }

$app = $null
$deck = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    foreach ($openDeck in $app.Presentations) {
        if ($openDeck.FullName -eq $source) {
            throw 'The input is open in PowerPoint. Render a separate copy of the saved file.'
        }
    }

    $previousSecurity = $app.AutomationSecurity
    try {
        $app.AutomationSecurity = 3
        $deck = $app.Presentations.Open($source, -1, 0, 0)
    }
    finally {
        $app.AutomationSecurity = $previousSecurity
    }
    if ($numbers[-1] -gt $deck.Slides.Count) { throw 'A slide number exceeds the slide count.' }
    New-Item -ItemType Directory -Path $destination | Out-Null
    $height = [int][Math]::Round($Width * $deck.PageSetup.SlideHeight / $deck.PageSetup.SlideWidth)
    foreach ($number in $numbers) {
        $path = Join-Path $destination ('slide_{0:D3}.png' -f $number)
        if (Test-Path -LiteralPath $path) { throw "Output already exists: $path" }
        $deck.Slides.Item($number).Export($path, 'PNG', $Width, $height)
        if (!(Test-Path -LiteralPath $path)) { throw "Missing render: $number" }
    }
    @{ output = $destination; slides = $numbers; width = $Width; height = $height } | ConvertTo-Json -Compress
}
finally {
    if ($null -ne $deck) {
        $deck.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($deck)
    }
    # An active PowerPoint process can also hold the user's presentations.
    if ($null -ne $app) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}
