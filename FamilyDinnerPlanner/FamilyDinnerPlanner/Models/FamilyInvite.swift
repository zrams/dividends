import Foundation

struct FamilyInvite: Identifiable, Hashable {
    let id: String
    let code: String
    let email: String
    let inviteLink: String
    let createdAt: Date?
    let claimedAt: Date?

    var isClaimed: Bool {
        claimedAt != nil
    }
}
