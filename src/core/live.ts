/** A block of lines redrawn in place on an interactive terminal, for progress panels that must stay off stdout. */
export interface LiveRegion {
  /** Use ANSI colours (off when NO_COLOR is set). */
  readonly color: boolean;
  /** Replace the lines drawn by the previous call with these. */
  render(lines: string[]): void;
}

const ANSI = /\x1b\[[0-9;]*m/y;

/** Cut a line to `width` visible characters, keeping colour codes intact, so each line takes one terminal row. */
export function fitLine(line: string, width: number): string {
  let out = '';
  let visible = 0;
  let styled = false;
  for (let i = 0; i < line.length;) {
    ANSI.lastIndex = i;
    const code = ANSI.exec(line);
    if (code) {
      out += code[0];
      styled = true;
      i += code[0].length;
      continue;
    }
    if (visible < width) {
      out += line[i];
      visible++;
    }
    i++;
  }
  return styled ? `${out}\x1b[0m` : out;
}

/** Live region on `stream` (stderr): moves the cursor back over the last frame and clears it before drawing. */
export function terminalRegion(stream: NodeJS.WriteStream): LiveRegion {
  let drawn = 0;
  return {
    color: !process.env.NO_COLOR,
    render(lines) {
      const width = Math.max(20, (stream.columns || 100) - 1);
      const back = drawn ? `\x1b[${drawn}F` : '';
      stream.write(`${back}\x1b[0J${lines.map((line) => fitLine(line, width)).join('\n')}\n`);
      drawn = lines.length;
    },
  };
}
