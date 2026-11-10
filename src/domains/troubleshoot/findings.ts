import type { AnyCommand, Permissions } from '../../core/types.ts';
import { isSettledError } from '../../core/util.ts';

export type Severity = 'high' | 'medium' | 'low' | 'info';

/** One thing the admin should know, in plain language, with the command that acts on it. */
export interface Finding {
  severity: Severity;
  area: string;
  issue: string;
  detail?: string;
  /** Exact, runnable command (no prose) to investigate or fix it; agents may run it verbatim. */
  next?: string;
}

const RANK: Record<Severity, number> = { high: 0, medium: 1, low: 2, info: 3 };

export function sortFindings(findings: Finding[]): Finding[] {
  return [...findings].sort((a, b) => RANK[a.severity] - RANK[b.severity]);
}

/** One sentence for the top of a report: what is wrong, or that nothing is. */
export function verdict(subject: string, findings: Finding[], unchecked: string[]): string {
  const problems = findings.filter((f) => f.severity === 'high' || f.severity === 'medium');
  const gap = unchecked.length ? ` (could not check: ${unchecked.join(', ')})` : '';
  if (!problems.length) return `No problems found for ${subject}${gap}.`;
  const first = problems[0] as Finding;
  const more = problems.length > 1 ? ` and ${problems.length - 1} more` : '';
  return `${problems.length} problem${problems.length === 1 ? '' : 's'} for ${subject}: ${first.issue}${more}${gap}.`;
}

/** Sections that failed (usually a missing permission or licence), so the agent knows what was not checked. */
export function unavailable(sections: Record<string, unknown>): { section: string; code: string; message: string }[] {
  const out: { section: string; code: string; message: string }[] = [];
  for (const [section, value] of Object.entries(sections)) {
    if (isSettledError(value)) out.push({ section, ...value.error });
  }
  return out;
}

/**
 * A composite view needs what each command it composes needs. Taking the union from those commands' own
 * declarations keeps `m365 permissions` and the login scope set accurate as they change.
 */
export function permissionsOf(sources: AnyCommand[], ids: string[], roles: string[]): Permissions {
  const byId: Record<string, AnyCommand> = {};
  for (const command of sources) byId[`${command.domain} ${command.name}`] = command;
  const delegated = new Set<string>();
  const application = new Set<string>();
  const defender = new Set<string>();
  let exchange = false;
  for (const id of ids) {
    const command = byId[id];
    if (!command) throw new Error(`troubleshoot: unknown source command "${id}"`);
    for (const p of command.permissions.delegated ?? []) delegated.add(p);
    for (const p of command.permissions.application ?? []) application.add(p);
    for (const p of command.permissions.defender ?? []) defender.add(p);
    exchange ||= command.permissions.exchange === true;
  }
  return {
    delegated: [...delegated].sort(),
    application: [...application].sort(),
    roles,
    ...(exchange ? { exchange } : {}),
    ...(defender.size ? { defender: [...defender].sort() } : {}),
  };
}
