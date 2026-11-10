import type { AnyCommand } from '../../core/types.ts';
import { commands as device } from './device.ts';
import { commands as mail } from './mail.ts';
import { commands as tenant } from './tenant.ts';
import { commands as user } from './user.ts';

export const commands: AnyCommand[] = [...user, ...device, ...mail, ...tenant];
