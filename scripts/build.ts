/**
 * Compile `m365` into standalone executables (no Bun or Node needed on the target machine).
 *
 *   bun run build              all targets into dist/
 *   bun run build --install    this machine only, then copy to ~/.local/bin (or %LOCALAPPDATA%\m365 on Windows)
 */
import { $ } from 'bun';
import { chmodSync, copyFileSync, existsSync, mkdirSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

const TARGETS = ['bun-darwin-arm64', 'bun-darwin-x64', 'bun-linux-x64', 'bun-linux-arm64', 'bun-windows-x64'];
const root = join(import.meta.dir, '..');
const install = process.argv.includes('--install');

const hostTarget = `bun-${process.platform === 'win32' ? 'windows' : process.platform}-${process.arch}`;
const targets = install ? [hostTarget] : TARGETS;
mkdirSync(join(root, 'dist'), { recursive: true });

for (const target of targets) {
  const outfile = join(root, 'dist', `m365-${target.slice(4)}${target.includes('windows') ? '.exe' : ''}`);
  await $`bun build ${join(root, 'src/cli.ts')} --compile --minify --target=${target} --outfile ${outfile}`.quiet();
  console.log(`built ${outfile}`);
  if (install) {
    const dir =
      process.platform === 'win32'
        ? join(process.env.LOCALAPPDATA ?? homedir(), 'm365')
        : join(homedir(), '.local', 'bin');
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
    const dest = join(dir, process.platform === 'win32' ? 'm365.exe' : 'm365');
    copyFileSync(outfile, dest);
    if (process.platform !== 'win32') chmodSync(dest, 0o755);
    console.log(`installed ${dest}`);
    if (!(process.env.PATH ?? '').split(process.platform === 'win32' ? ';' : ':').includes(dir)) {
      console.log(`\n${dir} is not on your PATH yet. To run \`m365\` from anywhere:`);
      if (process.platform === 'win32') {
        console.log(`  setx PATH "%PATH%;${dir}"   (then open a new terminal)`);
      } else {
        const shell = (process.env.SHELL ?? '').split('/').pop();
        const rc =
          shell === 'zsh' ? '~/.zshrc' : shell === 'bash' ? '~/.bashrc' : shell === 'fish' ? null : '~/.profile';
        console.log(
          rc
            ? `  echo 'export PATH="$HOME/.local/bin:$PATH"' >> ${rc} && source ${rc}`
            : '  fish_add_path ~/.local/bin',
        );
      }
    }
  }
}
