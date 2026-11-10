import type { AnyCommand } from '../../core/types.ts';
import { commands as actions } from './actions.ts';
import { commands as assignments } from './assignments.ts';
import { commands as devices } from './devices.ts';
import { commands as policies } from './policies.ts';

export { type ManagedDevice, platformOf, resolveManagedDevice, shapeManagedDevice } from './common.ts';
export {
  complianceSummary,
  devicePolicyStates,
  type PolicyState,
  type PolicyStateView,
  shapePolicyStates,
} from './devices.ts';
export { type Assignment, listAssignments, shapeAssignment, type Subject } from './assignments.ts';

export const commands: AnyCommand[] = [...devices, ...actions, ...policies, ...assignments];
