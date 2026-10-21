param(
  [Parameter(Mandatory = $true)][string]$OutputDirectory,
  [ValidateSet('KWPP.Application', 'PowerPoint.Application', 'PowerPoint.Application.16')]
  [string]$ApplicationProgId = 'KWPP.Application',
  [string]$ApplicationExecutablePath,
  [ValidateSet('Iterated', 'Motion', 'Word', 'CommonPresets', 'Emphasis', 'Transitions')][string]$Suite = 'Iterated'
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
    @{ Name = 'FlyExit'; Id = 2; Exit = $true },
    @{ Name = 'PathCircle'; Id = 86; ByCharacter = $false },
    @{ Name = 'PathWave'; Id = 132; ByCharacter = $false },
    @{ Name = 'PathArcDown'; Id = 122; ByCharacter = $false },
    @{ Name = 'PathBounceLeft'; Id = 126; ByCharacter = $false },
    @{ Name = 'PathDecayingWave'; Id = 145; ByCharacter = $false },
    @{ Name = 'PathFigure8Four'; Id = 113; ByCharacter = $false },
    @{ Name = 'PathHeart'; Id = 94; ByCharacter = $false },
    @{ Name = 'PathLoopdeLoop'; Id = 109; ByCharacter = $false },
    @{ Name = 'PathSineWave'; Id = 125; ByCharacter = $false },
    @{ Name = 'PathSpiralLeft'; Id = 140; ByCharacter = $false },
    @{ Name = 'PathSpring'; Id = 138; ByCharacter = $false },
    @{ Name = 'PathStairsDown'; Id = 147; ByCharacter = $false },
    @{ Name = 'PathTurnUpRight'; Id = 142; ByCharacter = $false },
    @{ Name = 'PathZigzag'; Id = 123; ByCharacter = $false },
    @{ Name = 'PathCircleAbsolute'; Id = 86; ByCharacter = $false; Relative = $false }
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
    # Microsoft MsoAnimEffect enumeration: msoAnimEffectWedge = 20.
    @{ Name = 'Wedge'; Id = 20; Text = 'Wedge'; ByCharacter = $false },
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
} elseif ($Suite -eq 'Transitions') {
  # PpEntryEffect constants follow Microsoft's PowerPoint object model.
  $cases = @(
    @{ Name = 'Fade'; EntryEffect = 1793 },
    @{ Name = 'Dissolve'; EntryEffect = 1537 },
    @{ Name = 'WipeRight'; EntryEffect = 2819 },
    @{ Name = 'BoxIn'; EntryEffect = 3074 },
    @{ Name = 'CheckerboardAcross'; EntryEffect = 1025 },
    @{ Name = 'CoverUp'; EntryEffect = 1282 },
    @{ Name = 'UncoverLeft'; EntryEffect = 2049 },
    @{ Name = 'PushLeft'; EntryEffect = 3853 },
    @{ Name = 'SplitHorizontalOut'; EntryEffect = 3585 },
    @{ Name = 'SplitVerticalIn'; EntryEffect = 3588 },
    @{ Name = 'CircleOut'; EntryEffect = 3845 },
    @{ Name = 'DiamondOut'; EntryEffect = 3846 },
    @{ Name = 'PlusOut'; EntryEffect = 3851 },
    @{ Name = 'Wedge'; EntryEffect = 3856 },
    @{ Name = 'BlindsVertical'; EntryEffect = 770 },
    @{ Name = 'CombHorizontal'; EntryEffect = 3847 },
    @{ Name = 'RandomBarsVertical'; EntryEffect = 2306 },
    @{ Name = 'Wheel8Spokes'; EntryEffect = 3861 },
    @{ Name = 'StripsDownLeft'; EntryEffect = 2563 },
    @{ Name = 'DoorsVertical'; EntryEffect = 3884 },
    @{ Name = 'WindowHorizontal'; EntryEffect = 3887 },
    @{ Name = 'Honeycomb'; EntryEffect = 3898 },
    @{ Name = 'FlyThroughIn'; EntryEffect = 3890 },
    @{ Name = 'ConveyorRight'; EntryEffect = 3883 }
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
  $transitionBackgrounds = @(
    'FECACA', 'BFDBFE', 'BBF7D0', 'FDE68A', 'DCC8F9', 'A5F3FC', 'FBCFE8', 'C7D2FE',
    'FED7AA', 'A7F3D0', 'E9D5FF', 'FFA8A8', '93C5FD', '86EFAC', 'FCD34D', 'C4B5FD',
    '67E8F9', 'F9A8D4', 'A5B4FC', 'FDBA74', '6EE7B7', 'D8B4FE', '7DD3FC', 'FAA5A5'
  )
  foreach ($case in $cases) {
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    if ($Suite -eq 'Transitions') {
      $backgroundHex = $transitionBackgrounds[($presentation.Slides.Count - 1) % $transitionBackgrounds.Count]
      $backgroundOleRgb = $backgroundHex.Substring(4, 2) + $backgroundHex.Substring(2, 2) + $backgroundHex.Substring(0, 2)
      $slide.Background.Fill.ForeColor.RGB = [Convert]::ToInt32($backgroundOleRgb, 16)
    } else {
      $slide.Background.Fill.ForeColor.RGB = 16777215
    }
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
    if ($Suite -eq 'Transitions') {
      try {
        $slide.SlideShowTransition.EntryEffect = $case.EntryEffect
        $slide.SlideShowTransition.Speed = 2
        $case.Supported = $true
      } catch {
        $case.Supported = $false
        $case.SupportError = $_.Exception.Message
      }
      continue
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
    if ($case.ContainsKey('Relative')) {
      $effect.EffectParameters.Relative = if ($case.Relative) { -1 } else { 0 }
      $case.RelativeApplied = [int]$effect.EffectParameters.Relative
    }
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
  $presentation.Saved = -1
  $presentation.Close()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
  $presentation = $null
  if ($Suite -eq 'Transitions') {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($fixturePath)
    try {
      for ($index = 0; $index -lt $cases.Count; $index++) {
        $case = $cases[$index]
        $case.Assigned = $case.Supported -eq $true
        $entry = $archive.GetEntry(('ppt/slides/slide{0}.xml' -f ($index + 1)))
        if ($null -eq $entry) {
          throw "Exported transition fixture is missing slide $($index + 1)."
        }
        $reader = [IO.StreamReader]::new($entry.Open())
        try {
          $slideXml = $reader.ReadToEnd()
        } finally {
          $reader.Dispose()
        }
        $document = [Xml.XmlDocument]::new()
        $document.XmlResolver = $null
        $document.LoadXml($slideXml)
        $namespaces = [Xml.XmlNamespaceManager]::new($document.NameTable)
        $namespaces.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
        $transitions = $document.SelectNodes('//p:transition[*]', $namespaces)
        $case.Exported = $transitions.Count -gt 0
        $case.Supported = $case.Assigned -and $case.Exported
        if ($case.Exported) {
          $case.ExportedEffects = @($transitions | ForEach-Object { $_.FirstChild.LocalName })
          $case.ExportedEffect = $case.ExportedEffects[0]
          $case.ExportedTransitionXml = @($transitions | ForEach-Object { $_.OuterXml })
        } elseif ($case.Assigned) {
          $case.SupportError = 'WPS accepted the COM assignment but exported no transition effect child in the slide XML.'
        }
      }
    } finally {
      $archive.Dispose()
    }
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
