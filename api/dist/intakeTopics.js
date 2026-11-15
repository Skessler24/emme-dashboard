"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.TOPICS = void 0;
exports.getTopic = getTopic;
const billing_1 = require("./topics/billing");
const onboarding_1 = require("./topics/onboarding");
const referrals_1 = require("./topics/referrals");
const staffInformation_1 = require("./topics/staffInformation");
const staffExpectations_1 = require("./topics/staffExpectations");
const reflectiveSupervision_1 = require("./topics/reflectiveSupervision");
/** Every tile that has a "setup questions" tab. The id is the web address segment: /api/intake/{id}/... */
exports.TOPICS = [
    billing_1.billing,
    onboarding_1.onboarding,
    referrals_1.referrals,
    staffInformation_1.staffInformation,
    staffExpectations_1.staffExpectations,
    reflectiveSupervision_1.reflectiveSupervision,
];
function getTopic(id) {
    return exports.TOPICS.find(function (topic) { return topic.id === id; });
}
