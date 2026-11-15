import { TopicDef } from './intakeTypes';
import { billing } from './topics/billing';
import { onboarding } from './topics/onboarding';
import { referrals } from './topics/referrals';
import { staffInformation } from './topics/staffInformation';
import { staffExpectations } from './topics/staffExpectations';
import { reflectiveSupervision } from './topics/reflectiveSupervision';

/** Every tile that has a "setup questions" tab. The id is the web address segment: /api/intake/{id}/... */
export const TOPICS: TopicDef[] = [
  billing,
  onboarding,
  referrals,
  staffInformation,
  staffExpectations,
  reflectiveSupervision,
];

export function getTopic(id: string): TopicDef | undefined {
  return TOPICS.find(function (topic) { return topic.id === id; });
}
