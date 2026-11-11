#!/bin/sh
# Install tenantctl on macOS or Linux from GitHub releases.
#
#   curl -fsSL https://raw.githubusercontent.com/willstarnet/tenantctl/main/scripts/install.sh | sh
#
# Environment:
#   TENANTCTL_VERSION      release to install, e.g. 1.2.0 or v1.2.0 (default: latest)
#   TENANTCTL_INSTALL_DIR  where to put the binary (default: ~/.local/bin)
#   TENANTCTL_REPO         GitHub owner/repo to download from (default: willstarnet/tenantctl)
#   TENANTCTL_BASE_URL     download archives and SHA256SUMS from here instead of GitHub (testing, mirrors)
set -eu

REPO=${TENANTCTL_REPO:-willstarnet/tenantctl}

say() { printf '%s\n' "$*"; }
die() {
  printf 'tenantctl install: %s\n' "$*" >&2
  exit 1
}

download() { # url dest
  if command -v curl >/dev/null 2>&1; then
    curl -fsSL --retry 3 -o "$2" "$1" || die "download failed: $1"
  elif command -v wget >/dev/null 2>&1; then
    wget -q -O "$2" "$1" || die "download failed: $1"
  else
    die 'curl or wget is required'
  fi
}

sha256() { # file
  if command -v sha256sum >/dev/null 2>&1; then
    sha256sum "$1" | cut -d ' ' -f 1
  elif command -v shasum >/dev/null 2>&1; then
    shasum -a 256 "$1" | cut -d ' ' -f 1
  elif command -v openssl >/dev/null 2>&1; then
    openssl dgst -sha256 -r "$1" | cut -d ' ' -f 1
  else
    die 'sha256sum, shasum or openssl is required to verify the download'
  fi
}

is_musl() {
  [ -f /etc/alpine-release ] && return 0
  ldd --version 2>&1 | grep -qi musl
}

target() {
  case "$(uname -s)" in
    Darwin) os=darwin ;;
    Linux) os=linux ;;
    MINGW* | MSYS* | CYGWIN*) die 'on Windows run: irm https://raw.githubusercontent.com/'"$REPO"'/main/scripts/install.ps1 | iex' ;;
    *) die "unsupported OS: $(uname -s)" ;;
  esac
  case "$(uname -m)" in
    x86_64 | amd64) arch=x64 ;;
    arm64 | aarch64) arch=arm64 ;;
    *) die "unsupported CPU architecture: $(uname -m)" ;;
  esac
  # A shell running under Rosetta reports x86_64 on Apple silicon; prefer the native build.
  if [ "$os" = darwin ] && [ "$arch" = x64 ] && [ "$(sysctl -n sysctl.proc_translated 2>/dev/null || echo 0)" = 1 ]; then
    arch=arm64
  fi
  suffix="$os-$arch"
  if [ "$os" = linux ]; then
    if is_musl; then suffix="$suffix-musl"; fi
    # x64 CPUs without AVX2 need the baseline build ("Illegal instruction" otherwise).
    if [ "$arch" = x64 ] && [ -r /proc/cpuinfo ] && ! grep -qw avx2 /proc/cpuinfo; then suffix="$suffix-baseline"; fi
  fi
  printf '%s' "$suffix"
}

main() {
  suffix=$(target)
  asset="tenantctl-$suffix.tar.gz"
  version=${TENANTCTL_VERSION:-latest}
  if [ -n "${TENANTCTL_BASE_URL:-}" ]; then
    base=${TENANTCTL_BASE_URL%/}
  elif [ "$version" = latest ]; then
    base="https://github.com/$REPO/releases/latest/download"
  else
    base="https://github.com/$REPO/releases/download/v${version#v}"
  fi
  dir=${TENANTCTL_INSTALL_DIR:-$HOME/.local/bin}

  tmp=$(mktemp -d 2>/dev/null || mktemp -d -t tenantctl)
  trap 'rm -rf "$tmp"' EXIT
  trap 'exit 130' INT TERM

  say "Downloading $asset ($version)"
  download "$base/$asset" "$tmp/$asset"
  download "$base/SHA256SUMS" "$tmp/SHA256SUMS"

  expected=$(awk -v f="$asset" '$2 == f || $2 == "*" f { print $1 }' "$tmp/SHA256SUMS")
  [ -n "$expected" ] || die "$asset is not listed in SHA256SUMS"
  actual=$(sha256 "$tmp/$asset")
  [ "$expected" = "$actual" ] || die "checksum mismatch for $asset (expected $expected, got $actual)"

  tar -xzf "$tmp/$asset" -C "$tmp" tenantctl
  mkdir -p "$dir"
  # Copy next to the destination then rename, so a running tenantctl is replaced atomically.
  cp "$tmp/tenantctl" "$dir/.tenantctl.new"
  chmod 755 "$dir/.tenantctl.new"
  mv -f "$dir/.tenantctl.new" "$dir/tenantctl"
  say "Installed $dir/tenantctl ($("$dir/tenantctl" --version))"

  case ":$PATH:" in
    *":$dir:"*) ;;
    *)
      say ""
      say "$dir is not on your PATH yet. To run \`tenantctl\` from anywhere:"
      case "${SHELL:-}" in
        */zsh) say "  echo 'export PATH=\"$dir:\$PATH\"' >> ~/.zshrc && . ~/.zshrc" ;;
        */bash) say "  echo 'export PATH=\"$dir:\$PATH\"' >> ~/.bashrc && . ~/.bashrc" ;;
        */fish) say "  fish_add_path $dir" ;;
        *) say "  echo 'export PATH=\"$dir:\$PATH\"' >> ~/.profile && . ~/.profile" ;;
      esac
      ;;
  esac
}

# Everything runs from main so a truncated `curl | sh` download cannot execute a partial script.
main "$@"
