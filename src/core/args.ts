import { CliError } from './errors.ts';
import type { ArgSpec } from './types.ts';

export function kebab(name: string): string {
  return name.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`);
}

function coerce(spec: ArgSpec, raw: string): unknown {
  switch (spec.type) {
    case 'number': {
      const value = Number(raw);
      if (raw.trim() === '' || !Number.isFinite(value)) {
        throw new CliError('USAGE', `--${kebab(spec.name)} expects a number, got "${raw}"`);
      }
      return value;
    }
    case 'boolean':
      if (/^(true|1|yes)$/i.test(raw)) return true;
      if (/^(false|0|no)$/i.test(raw)) return false;
      throw new CliError('USAGE', `--${kebab(spec.name)} expects true or false, got "${raw}"`);
    case 'string[]':
      if (spec.noSplit) return [raw];
      return raw
        .split(',')
        .map((part) => part.trim())
        .filter(Boolean);
    default:
      return raw;
  }
}

function checkChoice(spec: ArgSpec, value: unknown): void {
  if (!spec.choices) return;
  const values = Array.isArray(value) ? value : [value];
  for (const entry of values) {
    if (!spec.choices.includes(String(entry))) {
      throw new CliError('USAGE', `--${kebab(spec.name)} must be one of: ${spec.choices.join(', ')} (got "${entry}")`);
    }
  }
}

/** Parse command tokens (global flags already removed) against the command's arg specs. */
export function parseArgs(specs: readonly ArgSpec[], tokens: string[]): Record<string, unknown> {
  const byFlag = new Map<string, ArgSpec>();
  for (const spec of specs) {
    byFlag.set(kebab(spec.name), spec);
    if (spec.short) byFlag.set(spec.short, spec);
  }
  const values: Record<string, unknown> = {};
  const positionals: string[] = [];
  const assign = (spec: ArgSpec, raw: string) => {
    const value = coerce(spec, raw);
    if (spec.type === 'string[]') {
      values[spec.name] = [...((values[spec.name] as string[] | undefined) ?? []), ...(value as string[])];
    } else {
      values[spec.name] = value;
    }
  };

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i] as string;
    if (token === '--') {
      positionals.push(...tokens.slice(i + 1));
      break;
    }
    const long = token.match(/^--([^=]+)(?:=(.*))?$/s);
    const short = token.match(/^-([a-zA-Z])$/);
    if (!long && !short) {
      positionals.push(token);
      continue;
    }
    const flag = long ? (long[1] as string) : (short?.[1] as string);
    let inline = long?.[2];
    const spec = byFlag.get(flag);
    if (!spec && flag.startsWith('no-')) {
      const negated = byFlag.get(flag.slice(3));
      if (negated?.type === 'boolean' && inline === undefined) {
        values[negated.name] = false;
        continue;
      }
    }
    if (!spec) {
      throw new CliError('USAGE', `Unknown option ${token.split('=')[0]}`, {
        hint: 'Run the command with --help to see its options.',
      });
    }
    if (spec.type === 'boolean' && inline === undefined) {
      const next = tokens[i + 1];
      if (next !== undefined && /^(true|false)$/i.test(next)) {
        inline = next;
        i++;
      } else {
        values[spec.name] = true;
        continue;
      }
    }
    if (inline === undefined) {
      const next = tokens[i + 1];
      if (next === undefined || (next.startsWith('--') && next.length > 2)) {
        throw new CliError('USAGE', `${token} needs a value`);
      }
      inline = next;
      i++;
    }
    assign(spec, inline);
  }

  const positionalSpecs = specs.filter((spec) => spec.positional);
  for (const spec of positionalSpecs) {
    if (!positionals.length) break;
    if (values[spec.name] !== undefined) continue;
    if (spec.type === 'string[]') {
      for (const raw of positionals.splice(0)) assign(spec, raw);
    } else {
      assign(spec, positionals.shift() as string);
    }
  }
  if (positionals.length) {
    throw new CliError('USAGE', `Unexpected argument${positionals.length > 1 ? 's' : ''}: ${positionals.join(' ')}`, {
      hint: 'Quote values that contain spaces, e.g. "Jane Doe".',
    });
  }

  for (const spec of specs) {
    if (values[spec.name] === undefined) {
      if (spec.default !== undefined) values[spec.name] = spec.default;
      else if (spec.type === 'boolean') values[spec.name] = false;
      else if (spec.required) {
        throw new CliError(
          'USAGE',
          `Missing required ${spec.positional ? 'argument' : 'option'} --${kebab(spec.name)}`,
          {
            hint: spec.description,
          },
        );
      }
    }
    if (values[spec.name] !== undefined) checkChoice(spec, values[spec.name]);
  }
  return values;
}
