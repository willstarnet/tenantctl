import type { AnyCommand } from '../../core/types.ts';
import { commands as files } from './files.ts';
import { commands as onedrive } from './onedrive.ts';
import { commands as settings } from './settings.ts';
import { commands as sites } from './sites.ts';

export { shapePermission } from './files.ts';
export { resolveDriveItem, resolveSite } from './resolve.ts';

export const commands: AnyCommand[] = [...sites, ...files, ...onedrive, ...settings];
