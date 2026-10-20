param(
  [Parameter(Mandatory = $true)][string]$OutputDirectory,
  [ValidateSet('KWPP.Application', 'PowerPoint.Application', 'PowerPoint.Application.16')]
  [string]$ApplicationProgId = 'KWPP.Application',
  [string]$ApplicationExecutablePath,
  [ValidateSet('Iterated', 'Motion', 'Word', 'CommonPresets', 'Emphasis')][string]$Suite = 'Iterated'
)

$ErrorActionPreference = 'Stop'
$outputRoot = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $outputRoot) {
  if (-not (Get-Item -LiteralPath $outputRoot).PSIsContainer -or @(Get-ChildItem -LiteralPath $outputRoot -Force).Count -gt 0) {
    throw 'Use a new or empty output directory to preserve existing files.'
  }
} else {
  New-Item -ItemType Directory -Path $outputRoot | Out-Null
}

$initialOfficeProcesses = @(Get-Process wpp, POWERPNT -ErrorAction SilentlyContinue)
$launchedApplicationProcess = $null
$automation = $null
$presentation = $null
$cases = @(
  @{ Name = 'Fade'; Id = 10 },
  @{ Name = 'Zoom'; Id = 23 },
  @{ Name = 'Checkerboard'; Id = 5 },
  @{ Name = 'Blinds'; Id = 3 },
  @{ Name = 'Dissolve'; Id = 9 },
  @{ Name = 'Wheel'; Id = 21 },
  @{ Name = 'Box'; Id = 4 },
  @{ Name = 'RandomBars'; Id = 14 },
  @{ Name = 'Strips'; Id = 18 },
  @{ Name = 'Wipe'; Id = 22 }
)
if ($Suite -eq 'Motion') {
  $cases = @(
    @{ Name = 'Fly'; Id = 2 },
    @{ Name = 'Bounce'; Id = 26 },
    @{ Name = 'Spin'; Id = 61; ByCharacter = $false },
    @{ Name = 'GrowShrink'; Id = 59; ByCharacter = $false },
    @{ Name = 'GrowAndTurn'; Id = 31; ByCharacter = $false },
    @{ Name = 'Blast'; Id = 64; ByCharacter = $false },
    @{ Name = 'RandomEffects'; Id = 24; ByCharacter = $false },
    @{ Name = 'VerticalGrow'; Id = 81; ByCharacter = $false },
    @{ Name = 'Swivel'; Id = 19; ByCharacter = $false },
    @{ Name = 'Swish'; Id = 35 },
    @{ Name = 'FadedSwivel'; Id = 41 },
    @{ Name = 'FadeExit'; Id = 10; Exit = $true },
    @{ Name = 'ZoomExit'; Id = 23; Exit = $true },
    @{ Name = 'FlyExit'; Id = 2; Exit = $true }
  )
} elseif ($Suite -eq 'Word') {
  $cases = @(
    @{ Name = 'FadeByWord'; Id = 10; Text = 'ONE TWO'; ByCharacter = $false; ByWord = $true },
    @{ Name = 'WipeByWord'; Id = 22; Text = 'ONE TWO'; ByCharacter = $false; ByWord = $true },
    @{ Name = 'RandomBarsByWord'; Id = 14; Text = 'ONE TWO'; ByCharacter = $false; ByWord = $true },
    @{ Name = 'StripsByWord'; Id = 18; Text = 'ONE TWO'; ByCharacter = $false; ByWord = $true },
    # WPS exports this emphasis preset as p:animClr with an omitted from color.
    # Its OOXML default is black, and its destination is the theme accent2 color.
    @{ Name = 'ColorByWord'; Id = 56; PresetId = 3; PresetClass = 'emph'; Text = 'ONE TWO'; ByCharacter = $false; ByWord = $true }
  )
} elseif ($Suite -eq 'Emphasis') {
  # IDs and parameters follow Microsoft's PowerPoint object model.
  # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
  $cases = @(
    @{ Name = 'ChangeFontSize'; Id = 57; ByCharacter = $false },
    @{ Name = 'ChangeFont'; Id = 55; ByCharacter = $false },
    @{ Name = 'ChangeFontColor'; Id = 56; ByCharacter = $false },
    @{ Name = 'Transparency'; Id = 62; ByCharacter = $false },
    @{ Name = 'BoldFlash'; Id = 63; ByCharacter = $false },
    @{ Name = 'BoldReveal'; Id = 65; ByCharacter = $false },
    @{ Name = 'BrushOnUnderline'; Id = 67; ByCharacter = $false },
    @{ Name = 'FlashBulb'; Id = 75; ByCharacter = $false },
    @{ Name = 'ChangeFontStyle'; Id = 58; ByCharacter = $false },
    @{ Name = 'ChangeFillColor'; Id = 54; ByCharacter = $false },
    @{ Name = 'ChangeLineColor'; Id = 60; ByCharacter = $false },
    @{ Name = 'Spin'; Id = 61; ByCharacter = $false },
    @{ Name = 'GrowShrink'; Id = 59; ByCharacter = $false },
    @{ Name = 'Shimmer'; Id = 52; ByCharacter = $false },
    @{ Name = 'GrowWithColor'; Id = 77; ByCharacter = $false },
    @{ Name = 'Lighten'; Id = 78; ByCharacter = $false },
    @{ Name = 'Desaturate'; Id = 74; ByCharacter = $false }
  )
} elseif ($Suite -eq 'CommonPresets') {
  $cases = @(
    @{ Name = 'Crawl'; Id = 7; Text = 'Crawl'; ByCharacter = $false },
    @{ Name = 'Peek'; Id = 12; Text = 'Peek'; ByCharacter = $false },
    @{ Name = 'Stretch'; Id = 17; Text = 'Stretch'; ByCharacter = $false },
    @{ Name = 'Spiral'; Id = 15; Text = 'Spiral'; ByCharacter = $false },
    @{ Name = 'RandomEffects'; Id = 24; Text = 'RandomEffects'; ByCharacter = $false },
    @{ Name = 'GrowAndTurn'; Id = 31; Text = 'GrowAndTurn'; ByCharacter = $false },
    @{ Name = 'RiseUp'; Id = 34; Text = 'RiseUp'; ByCharacter = $false },
    @{ Name = 'Swish'; Id = 35; Text = 'Swish'; ByCharacter = $false },
    @{ Name = 'Ascend'; Id = 39; Text = 'Ascend'; ByCharacter = $false },
    @{ Name = 'FadedZoom'; Id = 48; Text = 'FadedZoom'; ByCharacter = $false },
    @{ Name = 'Glide'; Id = 49; Text = 'Glide'; ByCharacter = $false },
    @{ Name = 'Expand'; Id = 50; Text = 'Expand'; ByCharacter = $false },
    @{ Name = 'Flip'; Id = 51; Text = 'Flip'; ByCharacter = $false },
    @{ Name = 'Fold'; Id = 53; Text = 'Fold'; ByCharacter = $false },
    @{ Name = 'Blast'; Id = 64; Text = 'Blast'; ByCharacter = $false },
    @{ Name = 'Teeter'; Id = 80; Text = 'Teeter'; ByCharacter = $false },
    @{ Name = 'VerticalGrow'; Id = 81; Text = 'VerticalGrow'; ByCharacter = $false }
  )
}

try {
  if ($ApplicationExecutablePath) {
    $applicationExecutable = (Resolve-Path -LiteralPath $ApplicationExecutablePath).Path
    $launchedApplicationProcess = Start-Process -FilePath $applicationExecutable -ArgumentList '/automation' -WindowStyle Hidden -PassThru
    Start-Sleep -Seconds 3
  }
  $automation = New-Object -ComObject $ApplicationProgId
  if ($ApplicationExecutablePath) {
    $expectedApplicationPath = [IO.Path]::GetDirectoryName((Resolve-Path -LiteralPath $ApplicationExecutablePath).Path).TrimEnd('\')
    if (-not [string]::Equals($automation.Path.TrimEnd('\'), $expectedApplicationPath, [StringComparison]::OrdinalIgnoreCase)) {
      throw "Requested application executable '$expectedApplicationPath' but COM connected to '$($automation.Path)'."
    }
  }
  $presentation = $automation.Presentations.Add(0)
  $presentation.PageSetup.SlideWidth = 960
  $presentation.PageSetup.SlideHeight = 540
  foreach ($case in $cases) {
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = 16777215
    $shape = $slide.Shapes.AddTextbox(1, 120, 180, 720, 160)
    $byCharacter = -not $case.ContainsKey('ByCharacter') -or $case.ByCharacter
    $byWord = $case.ByWord -eq $true
    $shape.Name = $(if ($byCharacter -or $byWord) { 'Iterated ' } else { 'Whole ' }) + $case.Name
    $shape.TextFrame.MarginLeft = 0
    $shape.TextFrame.MarginRight = 0
    $shape.TextFrame.MarginTop = 0
    $shape.TextFrame.MarginBottom = 0
    $shape.TextFrame.WordWrap = 0
    $shape.TextFrame.TextRange.Text = $(if ($case.ContainsKey('Text')) { $case.Text } else { 'ABCD' })
    $shape.TextFrame.TextRange.Font.Name = 'Arial'
    $shape.TextFrame.TextRange.Font.Size = 96
    $shape.TextFrame.TextRange.Font.Color.RGB = 13395456
    if ($case.Name -in @('ChangeFillColor', 'GrowWithColor')) {
      $shape.Fill.Solid()
      $shape.Fill.ForeColor.RGB = 16777215
    }
    if ($case.Name -eq 'ChangeLineColor') {
      $shape.Line.Visible = -1
      $shape.Line.ForeColor.RGB = 0
      $shape.Line.Weight = 2
    }
    $sequence = $slide.TimeLine.MainSequence
    # Microsoft ConvertToTextUnitEffect: 1 means by character and 2 means by word.
    try {
      $effect = $sequence.AddEffect($shape, $case.Id, 0, 1)
    } catch {
      $case.Supported = $false
      $case.SupportError = $_.Exception.Message
      continue
    }
    $case.Supported = $true
    if ($case.Exit) { $effect.Exit = -1 }
    if ($byCharacter -or $byWord) { $effect = $sequence.ConvertToTextUnitEffect($effect, $(if ($byWord) { 2 } else { 1 })) }
    $effect.Timing.Duration = 2
    $effect.Timing.TriggerDelayTime = 0
  }
  $fixturePath = Join-Path $outputRoot ('office-' + $Suite.ToLowerInvariant() + '-effects.pptx')
  $presentation.SaveAs($fixturePath, 24)
  for ($index = 0; $index -lt $cases.Count; $index++) {
    $presentation.Slides.Item($index + 1).Export((Join-Path $outputRoot ($cases[$index].Name.ToLowerInvariant() + '-static.png')), 'PNG', 1280, 720)
  }
  # WPS can register PowerPoint.Application and report Microsoft's name/version.
  # Record executable identity as well, so it cannot be mistaken for PowerPoint evidence.
  $nativeProcesses = @(Get-Process wpp, POWERPNT -ErrorAction SilentlyContinue | ForEach-Object {
    $version = (Get-Item -LiteralPath $_.Path).VersionInfo
    @{ Id = $_.Id; Path = $_.Path; ProductVersion = $version.ProductVersion; Company = $version.CompanyName }
  })
  @{
    ProgId = $ApplicationProgId; ReportedApplicationName = $automation.Name
    CompatibilityVersion = $automation.Version; ApplicationPath = $automation.Path
    NativeProcesses = $nativeProcesses; Fixture = $fixturePath; Suite = $Suite; Cases = $cases
    ParameterMode = 'Native preset defaults'
  } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $outputRoot 'fixture-metadata.json')
  Write-Output $fixturePath
} finally {
  if ($null -ne $presentation) {
    $presentation.Saved = -1
    $presentation.Close()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
  }
  if ($null -ne $automation) {
    if ($initialOfficeProcesses.Count -eq 0 -and $automation.Presentations.Count -eq 0) { $automation.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($automation)
  }
  if ($null -ne $launchedApplicationProcess -and -not $launchedApplicationProcess.HasExited) {
    Stop-Process -Id $launchedApplicationProcess.Id -Force
  }
}
