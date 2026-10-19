param(
  [Parameter(Mandatory=$true)][string]$Config
)
# $Config: path to a JSON file with the action spec (Windows-readable path).
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$cfg = Get-Content -Raw -Path $Config | ConvertFrom-Json
$action = $cfg.action
$app = $null
$result = @{ ok = $false; action = $action }

function Get-PowerPointApp {
  try {
    return [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
  } catch {
    $a = New-Object -ComObject PowerPoint.Application
    return $a
  }
}

try {
  $app = Get-PowerPointApp
  $app.Visible = -1   # msoTrue (MsoTriState)

  $winPath = $null
  if ($cfg.path) { $winPath = [string]$cfg.path }

  function Find-Pres([string]$p) {
    foreach ($pr in $app.Presentations) {
      if ($pr.FullName -and $pr.FullName -ieq $p) { return $pr }
    }
    return $null
  }

  function Open-Pres([string]$p) {
    $found = Find-Pres $p
    if ($found -ne $null) { return $found }
    $pr = $app.Presentations.Open($p)  # ReadOnly, Untitled, WithWindow
    Start-Sleep -Milliseconds 400
    return $pr
  }

  switch ($action) {
    'open' {
      if (-not $winPath) { throw 'action open requires path' }
      $pres = Open-Pres $winPath
      $result.slideCount = $pres.Slides.Count
      $result.file = $pres.FullName
      $result.ok = $true
    }
    'show' {
      if (-not $winPath) { throw 'action show requires path' }
      $pres = Open-Pres $winPath
      $ss = $pres.SlideShowSettings
      $ss.Run()
      Start-Sleep -Milliseconds 700
      $result.slideCount = $pres.Slides.Count
      $result.running = $true
      if ($cfg.slide -and [int]$cfg.slide -ge 1) {
        $app.SlideShowWindows.Item(1).View.GotoSlide([int]$cfg.slide)
        $result.startSlide = [int]$cfg.slide
      }
      $result.ok = $true
    }
    'next' {
      if ($app.SlideShowWindows.Count -ge 1) {
        $app.SlideShowWindows.Item(1).View.Next()
        $result.ok = $true
      } else {
        $result.ok = $false
        $result.error = 'no slideshow running'
      }
    }
    'prev' {
      if ($app.SlideShowWindows.Count -ge 1) {
        $app.SlideShowWindows.Item(1).View.Previous()
        $result.ok = $true
      } else {
        $result.ok = $false
        $result.error = 'no slideshow running'
      }
    }
    'end' {
      if ($app.SlideShowWindows.Count -ge 1) {
        $app.SlideShowWindows.Item(1).View.Exit()
        $result.ok = $true
      } else {
        $result.ok = $true
        $result.error = 'no slideshow running'
      }
    }
    'export' {
      if (-not $winPath) { throw 'action export requires path' }
      $pres = Open-Pres $winPath
      $out = [string]$cfg.outdir
      New-Item -ItemType Directory -Force -Path $out | Out-Null
      $w = if ($null -ne $cfg.width) { [int]$cfg.width } else { 1280 }
      $h = if ($null -ne $cfg.height) { [int]$cfg.height } else { 720 }
      $pres.Export($out, 'PNG', $w, $h)
      $result.slideCount = $pres.Slides.Count
      $result.outdir = $out
      $result.ok = $true
    }
    'close' {
      if ($winPath) {
        $pres = Find-Pres $winPath
        if ($pres -ne $null) {
          $pres.Saved = -1
          $pres.Close()
          $result.ok = $true
        } else {
          $result.ok = $true
          $result.error = 'not open'
        }
      } else {
        foreach ($pr in @($app.Presentations)) {
          $pr.Saved = -1
          $pr.Close()
        }
        $result.ok = $true
      }
    }
    'refresh' {
      if (-not $winPath) { throw 'action refresh requires path' }
      $pres = Find-Pres $winPath
      if ($pres -ne $null) {
        $pres.Saved = -1
        $pres.Close()
        Start-Sleep -Milliseconds 500
      }
      $pres = Open-Pres $winPath
      $result.slideCount = $pres.Slides.Count
      $result.ok = $true
    }
    'info' {
      $result.ok = $true
      $result.presentations = @(foreach ($pr in $app.Presentations) { $pr.FullName })
      $result.slideShowRunning = ($app.SlideShowWindows.Count -ge 1)
      $result.slideCount = if ($app.Presentations.Count -ge 1) { $app.Presentations.Item(1).Slides.Count } else { 0 }
    }
    default { throw "unknown action: $action" }
  }

  $result | ConvertTo-Json -Compress
} catch {
  $r = @{ ok = $false; action = $action; error = $_.Exception.Message }
  $r | ConvertTo-Json -Compress
}
