import type { AnyCommand } from '../../core/types.ts';
import { commands as bulk } from './bulk.ts';
import { commands as directory } from './directory.ts';
import { commands as governance } from './governance.ts';
import { commands as users } from './users.ts';

export const commands: AnyCommand[] = [...users, ...bulk, ...directory, ...governance];
