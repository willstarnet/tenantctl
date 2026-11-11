# Install tenantctl on Windows from GitHub releases (Windows PowerShell 5.1 or PowerShell 7+).
#
#   irm https://raw.githubusercontent.com/willstarnet/tenantctl/main/scripts/install.ps1 | iex
#
# Environment:
#   TENANTCTL_VERSION      release to install, e.g. 1.2.0 or v1.2.0 (default: latest)
#   TENANTCTL_INSTALL_DIR  where to put tenantctl.exe (default: %LOCALAPPDATA%\tenantctl)
#   TENANTCTL_REPO         GitHub owner/repo to download from (default: willstarnet/tenantctl)
#   TENANTCTL_BASE_URL     download the zip and SHA256SUMS from here instead of GitHub (testing, mirrors)
#
# Runs inside the caller's session under `iex`, so it throws on failure instead of calling `exit`.

function Install-Tenantctl {
  $ErrorActionPreference = 'Stop'
  Set-StrictMode -Version 2

  $repo = if ($env:TENANTCTL_REPO) { $env:TENANTCTL_REPO } else { 'willstarnet/tenantctl' }
  $version = if ($env:TENANTCTL_VERSION) { $env:TENANTCTL_VERSION } else { 'latest' }
  if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    throw "install.ps1 is for Windows. On macOS/Linux run: curl -fsSL https://raw.githubusercontent.com/$repo/main/scripts/install.sh | sh"
  }
  $dir = if ($env:TENANTCTL_INSTALL_DIR) { $env:TENANTCTL_INSTALL_DIR } else { Join-Path $env:LOCALAPPDATA 'tenantctl' }

  # The registry reports the native architecture even when this PowerShell runs as x64 under ARM64 emulation.
  $nativeArch = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment').PROCESSOR_ARCHITECTURE
  switch ($nativeArch) {
    'ARM64' { $suffix = 'windows-arm64' }
    'AMD64' {
      # x64 CPUs without AVX2 need the baseline build. PF_AVX2_INSTRUCTIONS_AVAILABLE = 40.
      $avx2 = $false
      try {
        if (-not ('Tenantctl.Cpu' -as [type])) {
          Add-Type -Namespace Tenantctl -Name Cpu -MemberDefinition '[DllImport("kernel32.dll")] public static extern bool IsProcessorFeaturePresent(uint feature);'
        }
        $avx2 = [Tenantctl.Cpu]::IsProcessorFeaturePresent(40)
      } catch {
        $avx2 = $false
      }
      $suffix = if ($avx2) { 'windows-x64' } else { 'windows-x64-baseline' }
    }
    default { throw "Unsupported Windows architecture: $nativeArch (tenantctl ships x64 and ARM64 builds)" }
  }

  $asset = "tenantctl-$suffix.zip"
  if ($env:TENANTCTL_BASE_URL) {
    $base = $env:TENANTCTL_BASE_URL.TrimEnd('/')
  } elseif ($version -eq 'latest') {
    $base = "https://github.com/$repo/releases/latest/download"
  } else {
    $base = "https://github.com/$repo/releases/download/v$($version.TrimStart('v'))"
  }

  # Windows PowerShell 5.1 defaults to TLS 1.0/1.1, which GitHub rejects.
  [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

  $tmp = Join-Path ([IO.Path]::GetTempPath()) ("tenantctl-" + [Guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Path $tmp | Out-Null
  $progress = $ProgressPreference
  try {
    # The progress bar makes Invoke-WebRequest many times slower on Windows PowerShell 5.1.
    $ProgressPreference = 'SilentlyContinue'
    Write-Host "Downloading $asset ($version)"
    Invoke-WebRequest -UseBasicParsing -Uri "$base/$asset" -OutFile (Join-Path $tmp $asset)
    Invoke-WebRequest -UseBasicParsing -Uri "$base/SHA256SUMS" -OutFile (Join-Path $tmp 'SHA256SUMS')

    $expected = $null
    foreach ($line in Get-Content (Join-Path $tmp 'SHA256SUMS')) {
      $parts = $line.Trim() -split '\s+'
      if ($parts.Count -eq 2 -and $parts[1].TrimStart('*') -eq $asset) { $expected = $parts[0] }
    }
    if (-not $expected) { throw "$asset is not listed in SHA256SUMS" }
    $actual = (Get-FileHash -Algorithm SHA256 -Path (Join-Path $tmp $asset)).Hash
    if ($actual -ne $expected.ToUpperInvariant()) {
      throw "Checksum mismatch for $asset (expected $expected, got $actual)"
    }

    Expand-Archive -Path (Join-Path $tmp $asset) -DestinationPath (Join-Path $tmp 'out') -Force
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $exe = Join-Path $dir 'tenantctl.exe'
    # A running .exe cannot be overwritten but can be renamed, so move the old one aside first.
    if (Test-Path $exe) {
      Move-Item -Path $exe -Destination "$exe.old" -Force
    }
    Move-Item -Path (Join-Path $tmp 'out\tenantctl.exe') -Destination $exe -Force
    Remove-Item -Path "$exe.old" -Force -ErrorAction SilentlyContinue
  } finally {
    $ProgressPreference = $progress
    Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue
  }

  # Add the install dir to the USER Path. Read and write the raw registry value so entries like %USERPROFILE%\bin
  # keep their REG_EXPAND_SZ form ([Environment]::GetEnvironmentVariable would expand them and save them expanded).
  $key = (Get-Item 'HKCU:\').OpenSubKey('Environment', $true)
  try {
    $userPath = [string]$key.GetValue('Path', '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    $entries = @($userPath -split ';' | Where-Object { $_ })
    $present = $entries | Where-Object { [Environment]::ExpandEnvironmentVariables($_).TrimEnd('\') -ieq $dir.TrimEnd('\') }
    if (-not $present) {
      $key.SetValue('Path', (($entries + $dir) -join ';'), [Microsoft.Win32.RegistryValueKind]::ExpandString)
      # Setting any user variable through .NET broadcasts WM_SETTINGCHANGE so new terminals see the new Path.
      [Environment]::SetEnvironmentVariable('TENANTCTL_INSTALL_PATH_REFRESH', '1', 'User')
      [Environment]::SetEnvironmentVariable('TENANTCTL_INSTALL_PATH_REFRESH', $null, 'User')
      Write-Host "Added $dir to your user Path (open a new terminal to use it everywhere)."
    }
  } finally {
    $key.Close()
  }
  if (-not (($env:Path -split ';') | Where-Object { $_.TrimEnd('\') -ieq $dir.TrimEnd('\') })) {
    $env:Path = "$env:Path;$dir"
  }

  $installed = & $exe --version
  Write-Host "Installed $exe ($installed)"
}

Install-Tenantctl
