import type { AnyCommand } from '../../core/types.ts';
import { commands as mailbox } from './mailbox.ts';
import { commands as powershell } from './powershell.ts';
import { commands as rooms } from './rooms.ts';
import { commands as rules } from './rules.ts';
import { commands as trace } from './trace.ts';

export { shapeAutoReply } from './mailbox.ts';
export {
  aliasesOf,
  flattenMailboxPermissions,
  type MailboxAccess,
  type MailboxPermissionEntry,
  parseExoBytes,
  shapeQuarantineMessage,
} from './powershell.ts';
export { isScheduleFree } from './rooms.ts';
export {
  getInboxRules,
  getInboxRulesBatch,
  getMailboxForwarding,
  isInternal,
  listTenantDomains,
  type MailboxForwarding,
  shapeInboxRule,
  type ShapedInboxRule,
} from './rules.ts';
export { buildTraceFilter, shapeMailboxUsage } from './trace.ts';

export const commands: AnyCommand[] = [...mailbox, ...rules, ...powershell, ...rooms, ...trace];
