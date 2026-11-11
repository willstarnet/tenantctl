/**
 * Generate package-manager manifests for a release from dist/SHA256SUMS (written by `bun run build --archive`).
 *
 *   bun run release:manifests [version]     version defaults to package.json `version`
 *
 * Writes dist/manifests/:
 *   tenantctl.rb                                Homebrew formula (for a tap such as willstarnet/homebrew-tap)
 *   tenantctl.json                              Scoop manifest (for a bucket), with checkver/autoupdate
 *   <Id>.yaml, <Id>.installer.yaml,
 *   <Id>.locale.en-US.yaml                      winget multi-file manifest (schema 1.28.0) for microsoft/winget-pkgs
 *
 * Package managers can't probe the CPU before choosing a download, so x64 Linux (Homebrew) and x64 Windows (Scoop,
 * winget) use the `-baseline` build, which runs on every x64 CPU. The AVX2 build only speeds up Bun internals this
 * CLI barely touches (it waits on Graph). install.sh / install.ps1 do probe and pick the faster build when they can.
 */
import { mkdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import pkg from '../package.json' with { type: 'json' };

const REPO = process.env.TENANTCTL_REPO ?? 'willstarnet/tenantctl';
const OWNER = REPO.split('/')[0]!;
const WINGET_ID = `${OWNER}.tenantctl`;
const WINGET_SCHEMA = '1.28.0';
const SUMMARY = 'Microsoft 365 tenant administration CLI over Microsoft Graph';
const pkgLicense = 'license' in pkg && typeof pkg.license === 'string' ? pkg.license : undefined;
const LICENSE = pkgLicense ?? 'Proprietary';
const HOMEPAGE = `https://github.com/${REPO}`;

const root = join(import.meta.dir, '..');
const dist = join(root, 'dist');
const out = join(dist, 'manifests');
const version = (process.argv[2] ?? pkg.version).replace(/^v/, '');
if (!/^\d+\.\d+\.\d+([-+][0-9A-Za-z.-]+)?$/.test(version)) throw new Error(`not a semver version: ${version}`);

const sums = new Map<string, string>();
for (const line of readFileSync(join(dist, 'SHA256SUMS'), 'utf8').split('\n')) {
  const [hash, file] = line.trim().split(/\s+/);
  if (hash && file) sums.set(file.replace(/^\*/, ''), hash);
}

function asset(suffix: string): { url: string; sha256: string } {
  const file = `tenantctl-${suffix}${suffix.startsWith('windows') ? '.zip' : '.tar.gz'}`;
  const sha256 = sums.get(file);
  if (!sha256) throw new Error(`${file} is missing from dist/SHA256SUMS; run \`bun run build --archive\` first`);
  return { url: `${HOMEPAGE}/releases/download/v${version}/${file}`, sha256 };
}

/** YAML double-quoted scalar (JSON strings are valid YAML). */
const q = (value: string) => JSON.stringify(value);
/** Ruby double-quoted string literal for the plain ASCII values used here. */
const rb = (value: string) => JSON.stringify(value).replaceAll('#', '\\#');

mkdirSync(out, { recursive: true });

// Homebrew formula
{
  const block = (suffix: string, indent: string) => {
    const { url, sha256 } = asset(suffix);
    return `${indent}url ${rb(url)}\n${indent}sha256 ${rb(sha256)}`;
  };
  const formula = `class Tenantctl < Formula
  desc ${rb(SUMMARY)}
  homepage ${rb(HOMEPAGE)}
  version ${rb(version)}
${pkgLicense ? `  license ${rb(LICENSE)}\n` : ''}
  on_macos do
    on_arm do
${block('darwin-arm64', '      ')}
    end
    on_intel do
${block('darwin-x64', '      ')}
    end
  end

  on_linux do
    on_arm do
${block('linux-arm64', '      ')}
    end
    on_intel do
${block('linux-x64-baseline', '      ')}
    end
  end

  def install
    bin.install "tenantctl"
  end

  test do
    assert_match version.to_s, shell_output("#{bin}/tenantctl --version")
  end
end
`;
  await Bun.write(join(out, 'tenantctl.rb'), formula);
}

// Scoop manifest
{
  const x64 = asset('windows-x64-baseline');
  const arm64 = asset('windows-arm64');
  const download = `${HOMEPAGE}/releases/download/v$version`;
  const manifest = {
    version,
    description: SUMMARY,
    homepage: HOMEPAGE,
    license: LICENSE,
    architecture: {
      '64bit': { url: x64.url, hash: x64.sha256 },
      arm64: { url: arm64.url, hash: arm64.sha256 },
    },
    bin: 'tenantctl.exe',
    checkver: { github: HOMEPAGE },
    autoupdate: {
      architecture: {
        '64bit': { url: `${download}/tenantctl-windows-x64-baseline.zip` },
        arm64: { url: `${download}/tenantctl-windows-arm64.zip` },
      },
      // Scoop finds the line for the downloaded file name in the release's SHA256SUMS.
      hash: { url: '$baseurl/SHA256SUMS' },
    },
  };
  await Bun.write(join(out, 'tenantctl.json'), `${JSON.stringify(manifest, null, 2)}\n`);
}

// winget manifests (version, installer, defaultLocale)
{
  const header = (type: string) =>
    `# yaml-language-server: $schema=https://aka.ms/winget-manifest.${type}.${WINGET_SCHEMA}.schema.json\n\n` +
    `PackageIdentifier: ${q(WINGET_ID)}\nPackageVersion: ${q(version)}\n`;
  const footer = (type: string) => `ManifestType: ${type}\nManifestVersion: ${WINGET_SCHEMA}\n`;
  const installer = (arch: string, suffix: string) => {
    const { url, sha256 } = asset(suffix);
    return `  - Architecture: ${arch}\n    InstallerUrl: ${q(url)}\n    InstallerSha256: ${sha256.toUpperCase()}\n`;
  };

  await Bun.write(join(out, `${WINGET_ID}.yaml`), `${header('version')}DefaultLocale: en-US\n${footer('version')}`);
  await Bun.write(
    join(out, `${WINGET_ID}.installer.yaml`),
    `${header('installer')}InstallerType: zip
NestedInstallerType: portable
NestedInstallerFiles:
  - RelativeFilePath: tenantctl.exe
    PortableCommandAlias: tenantctl
Commands:
  - tenantctl
ReleaseDate: ${new Date().toISOString().slice(0, 10)}
Installers:
${installer('x64', 'windows-x64-baseline')}${installer('arm64', 'windows-arm64')}${footer('installer')}`,
  );
  await Bun.write(
    join(out, `${WINGET_ID}.locale.en-US.yaml`),
    `${header('defaultLocale')}PackageLocale: en-US
Publisher: ${q(OWNER)}
PublisherUrl: ${q(`https://github.com/${OWNER}`)}
PackageName: tenantctl
PackageUrl: ${q(HOMEPAGE)}
License: ${q(LICENSE)}
ShortDescription: ${q(SUMMARY)}
Description: ${q(pkg.description)}
Moniker: tenantctl
Tags:
  - microsoft-365
  - microsoft-graph
  - entra-id
  - intune
  - cli
ReleaseNotesUrl: ${q(`${HOMEPAGE}/releases/tag/v${version}`)}
${footer('defaultLocale')}`,
  );
}

console.log(`wrote ${out} for v${version}`);
