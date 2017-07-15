function Get-ChildItemColor {
<#
.Synopsis
  Returns childitems with colors by type.
  From http://poshcode.org/?show=878
.Description
  This function wraps Get-ChildItem and tries to output the results
  color-coded by type:
  Compressed - Red
  Directories - Blue
  Executables - Green
  Text Files - White
  Configs - Yellow
  Symlinks - Cyan
  Others - Default
.ReturnValue
  All objects returned by Get-ChildItem are passed down the pipeline
  unmodified.
.Notes
  NAME:      Get-ChildItemColor
  AUTHOR:    Tojo2000 <tojo2000@tojo2000.com>
  MODIFIED:  Karloch <cmilanf@hispamsx.org>
                * Change colors to match UNIX style
                * Added background color for directories
                * Added symlink detection support
#>

function Test-ReparsePoint([string]$path) {
  $file = Get-Item $path -Force -ea 0
  return [bool]($file.Attributes -band [IO.FileAttributes]::ReparsePoint)
}

  $regex_opts = ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase `
      -bor [System.Text.RegularExpressions.RegexOptions]::Compiled)
 
  $fore = $Host.UI.RawUI.ForegroundColor
  $back = $Host.UI.RawUI.BackgroundColor
  $compressed = New-Object System.Text.RegularExpressions.Regex(
      '\.(zip|tar|gz|rar|7z)$', $regex_opts)
  $executable = New-Object System.Text.RegularExpressions.Regex(
      '\.(exe|com|bat|cmd|py|pl|ps1|psm1|vbs|rb|reg|fsx)$', $regex_opts)
  $dll_pdb = New-Object System.Text.RegularExpressions.Regex(
      '\.(dll|pdb)$', $regex_opts)
  $configs = New-Object System.Text.RegularExpressions.Regex(
      '\.(config|conf|ini|cfg)$', $regex_opts)
  $text_files = New-Object System.Text.RegularExpressions.Regex(
      '\.(txt|cfg|conf|ini|csv|log)$', $regex_opts)
  $links = New-Object System.Text.RegularExpressions.Regex(
      '\.(lnk)$', $regex_opts)

  Invoke-Expression ("Get-ChildItem $args") |
    %{
      $c = $fore
      $b = $back
      $file = Get-Item $_.Name -Force -ea 0
      if ([bool]($file.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        $c = 'Cyan'
      } elseif ($_.GetType().Name -eq 'DirectoryInfo') {
        $c = 'Blue'
        #$b = 'Gray'
      } elseif ($compressed.IsMatch($_.Name)) {
        $c = 'Red'
      } elseif ($executable.IsMatch($_.Name)) {
        $c = 'Green'
      } elseif ($text_files.IsMatch($_.Name)) {
        $c = 'White'
      } elseif ($dll_pdb.IsMatch($_.Name)) {
        $c = 'DarkGreen'
      } elseif ($configs.IsMatch($_.Name)) {
        $c = 'Yellow'
      } elseif ($links.IsMatch($_.Name)) {
        $c = 'Cyan'
      }
      $Host.UI.RawUI.ForegroundColor = $c
      $Host.UI.RawUI.BackgroundColor = $b
      echo $_
      $Host.UI.RawUI.ForegroundColor = $fore
      $Host.UI.RawUI.BackgroundColor = $back
    }
}