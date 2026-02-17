import Foundation

struct MemberSubmissionHistoryItem: Identifiable, Hashable {
    let id: String
    let weekStart: Date
    let choiceNames: [String]
}
