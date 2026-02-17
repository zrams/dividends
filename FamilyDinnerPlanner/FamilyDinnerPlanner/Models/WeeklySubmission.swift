import Foundation

struct WeeklySubmission: Identifiable, Codable, Hashable {
    let id: String
    let userId: String
    let weekStart: Date
    let choices: [String]

    var asFirestoreData: [String: Any] {
        [
            "id": id,
            "userId": userId,
            "weekStart": weekStart,
            "choices": choices
        ]
    }
}
