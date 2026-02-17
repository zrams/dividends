import Foundation

struct DinnerIdea: Identifiable, Codable, Hashable {
    let id: String
    var name: String
    var description: String?
}
