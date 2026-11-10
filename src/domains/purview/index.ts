import type { AnyCommand } from '../../core/types.ts';
import { commands as audit } from './audit.ts';
import { commands as ediscovery } from './ediscovery.ts';

export { type AuditFilters, auditQueryBody, awaitAuditQuery, runAuditQuery, shapeRecord } from './audit.ts';
export { resolveCase } from './ediscovery.ts';

export const commands: AnyCommand[] = [...audit, ...ediscovery];
