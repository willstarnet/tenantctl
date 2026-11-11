/**
 * Compile `tenantctl` into standalone executables (no Bun or Node needed on the target machine).
 *
 *   bun run build              all targets into dist/tenantctl-<target>[.exe]
 *   bun run build --archive    all targets, then dist/tenantctl-<target>.tar.gz|.zip + dist/SHA256SUMS (release assets)
 *   bun run build --install    this machine only, then copy to ~/.local/bin (or %LOCALAPPDATA%\tenantctl on Windows)
 *
 * The version baked into the binary is package.json `version` (src/cli.ts imports it); CI sets it from the tag.
 */
import { $ } from 'bun';
import { chmodSync, copyFileSync, existsSync, mkdirSync, rmSync } from 'node:fs';
import { homedir } from 'node:os';
import { basename, join } from 'node:path';

/**
 * Every `--compile` target Bun 1.3.x documents. `-baseline` builds are for x64 CPUs without AVX2 (pre-2013 and some
 * low-end Atom/Celeron/Pentium chips); the installers pick them automatically. macOS x64 needs no baseline: every Mac
 * that runs a macOS release Bun supports has AVX2.
 */
const TARGETS = [
  'bun-darwin-arm64',
  'bun-darwin-x64',
  'bun-linux-x64',
  'bun-linux-x64-baseline',
  'bun-linux-arm64',
  'bun-linux-x64-musl',
  'bun-linux-x64-musl-baseline',
  'bun-linux-arm64-musl',
  'bun-windows-x64',
  'bun-windows-x64-baseline',
  'bun-windows-arm64',
];

const root = join(import.meta.dir, '..');
const dist = join(root, 'dist');
const install = process.argv.includes('--install');
const archive = process.argv.includes('--archive');

const hostTarget = `bun-${process.platform === 'win32' ? 'windows' : process.platform}-${process.arch}`;
const targets = install ? [hostTarget] : TARGETS;
mkdirSync(dist, { recursive: true });

const binaries: { suffix: string; outfile: string; windows: boolean }[] = [];
for (const target of targets) {
  const suffix = target.slice('bun-'.length);
  const windows = target.includes('windows');
  const outfile = join(dist, `tenantctl-${suffix}${windows ? '.exe' : ''}`);
  // --bytecode needs --format=esm here: the CLI is ESM (top-level await, JSON import attribute).
  await $`bun build ${join(root, 'src/cli.ts')} --compile --minify --bytecode --format=esm --target=${target} --outfile ${outfile}`.quiet();
  console.log(`built ${outfile}`);
  binaries.push({ suffix, outfile, windows });
}

if (archive) {
  // One archive per target holding a single `tenantctl` / `tenantctl.exe`, plus SHA256SUMS in `sha256sum` format.
  const stage = join(dist, 'stage');
  const sums: string[] = [];
  for (const { suffix, outfile, windows } of binaries) {
    const dir = join(stage, suffix);
    rmSync(dir, { recursive: true, force: true });
    mkdirSync(dir, { recursive: true });
    const exe = windows ? 'tenantctl.exe' : 'tenantctl';
    copyFileSync(outfile, join(dir, exe));
    chmodSync(join(dir, exe), 0o755);
    const file = join(dist, `tenantctl-${suffix}${windows ? '.zip' : '.tar.gz'}`);
    rmSync(file, { force: true });
    if (windows) await $`zip -q -X -j ${file} ${join(dir, exe)}`;
    // COPYFILE_DISABLE keeps macOS tar from adding AppleDouble `._tenantctl` entries.
    else await $`tar -czf ${file} -C ${dir} ${exe}`.env({ ...process.env, COPYFILE_DISABLE: '1' });
    const hash = new Bun.CryptoHasher('sha256').update(await Bun.file(file).arrayBuffer()).digest('hex');
    sums.push(`${hash}  ${basename(file)}`);
    console.log(`archived ${file}`);
  }
  rmSync(stage, { recursive: true, force: true });
  await Bun.write(join(dist, 'SHA256SUMS'), `${sums.join('\n')}\n`);
  console.log(`wrote ${join(dist, 'SHA256SUMS')}`);
}

if (install) {
  const outfile = binaries[0]!.outfile;
  const windows = process.platform === 'win32';
  const dir = windows ? join(process.env.LOCALAPPDATA ?? homedir(), 'tenantctl') : join(homedir(), '.local', 'bin');
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
  const dest = join(dir, windows ? 'tenantctl.exe' : 'tenantctl');
  copyFileSync(outfile, dest);
  if (!windows) chmodSync(dest, 0o755);
  console.log(`installed ${dest}`);
  if (!(process.env.PATH ?? '').split(windows ? ';' : ':').includes(dir)) {
    console.log(`\n${dir} is not on your PATH yet. To run \`tenantctl\` from anywhere:`);
    if (windows) {
      console.log(`  setx PATH "%PATH%;${dir}"   (then open a new terminal)`);
    } else {
      const shell = (process.env.SHELL ?? '').split('/').pop();
      const rc = shell === 'zsh' ? '~/.zshrc' : shell === 'bash' ? '~/.bashrc' : shell === 'fish' ? null : '~/.profile';
      console.log(
        rc ? `  echo 'export PATH="$HOME/.local/bin:$PATH"' >> ${rc} && source ${rc}` : '  fish_add_path ~/.local/bin',
      );
    }
  }
}
