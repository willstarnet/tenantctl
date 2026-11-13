import type { AnyCommand } from '../../core/types.ts';
import { commands as actions } from './actions.ts';
import { commands as assignments } from './assignments.ts';
import { commands as autopilot } from './autopilot.ts';
import { commands as devices } from './devices.ts';
import { commands as enrollment } from './enrollment.ts';
import { commands as filters } from './filters.ts';
import { commands as policies } from './policies.ts';

export { type ManagedDevice, platformOf, resolveManagedDevice, shapeManagedDevice } from './common.ts';
export {
  type ConfigProfileView,
  complianceSummary,
  deviceConfigurationReport,
  devicePolicyStates,
  groupReportRows,
  type PolicyStateView,
  shapePolicyStates,
} from './devices.ts';
export { type Assignment, listAssignments, shapeAssignment, type Subject } from './assignments.ts';

export const commands: AnyCommand[] = [
  ...devices,
  ...actions,
  ...policies,
  ...assignments,
  ...filters,
  ...autopilot,
  ...enrollment,
];
