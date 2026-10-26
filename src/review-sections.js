import {getSectionReview as contractSectionReview} from './contract-sections.js?v=20261026-opening-strategy';
import {getOpeningRuleReview} from './opening-rules.js?v=20261026-opening-strategy';

export function getSectionReview(task,factIds){
 return getOpeningRuleReview(task,factIds)||contractSectionReview(task,factIds);
}
