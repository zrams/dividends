import Foundation
import FirebaseFirestore
import Observation

@MainActor
@Observable
final class FamilyManagementViewModel {
    private let db = Firestore.firestore()
    private var membersListener: ListenerRegistration?
    private var invitesListener: ListenerRegistration?
    private var currentFamilyId: String?

    var members: [FamilyMember] = []
    var invites: [FamilyInvite] = []
    var inviteEmail = ""
    var latestInviteLink: String?
    var latestInviteCode: String?
    var isLoadingMembers = false
    var isCreatingInvite = false
    var errorMessage: String?

    func startListening(familyId: String) {
        guard !familyId.isEmpty else {
            errorMessage = "Family is not configured yet."
            return
        }

        if currentFamilyId == familyId, membersListener != nil, invitesListener != nil {
            return
        }

        stopListening()
        currentFamilyId = familyId
        isLoadingMembers = true
        errorMessage = nil

        membersListener = db
            .collection("users")
            .whereField("role", isEqualTo: "member")
            .whereField("familyId", isEqualTo: familyId)
            .addSnapshotListener { [weak self] snapshot, error in
                Task { @MainActor [weak self] in
                    guard let self else { return }

                    if let error {
                        self.errorMessage = error.localizedDescription
                        self.isLoadingMembers = false
                        return
                    }

                    self.members = (snapshot?.documents ?? []).map { document in
                        let data = document.data()
                        let email = (data["email"] as? String) ?? ""
                        let name = self.resolveName(
                            from: data["name"] as? String,
                            email: email,
                            fallback: document.documentID
                        )

                        return FamilyMember(
                            id: document.documentID,
                            name: name,
                            email: email.isEmpty ? "No email" : email,
                            familyId: data["familyId"] as? String ?? familyId
                        )
                    }
                    .sorted { lhs, rhs in
                        lhs.name.localizedCaseInsensitiveCompare(rhs.name) == .orderedAscending
                    }

                    self.isLoadingMembers = false
                }
            }

        invitesListener = db
            .collection("invites")
            .whereField("familyId", isEqualTo: familyId)
            .addSnapshotListener { [weak self] snapshot, error in
                Task { @MainActor [weak self] in
                    guard let self else { return }

                    if let error {
                        self.errorMessage = error.localizedDescription
                        return
                    }

                    self.invites = (snapshot?.documents ?? []).map { document in
                        let data = document.data()
                        let createdAt = (data["createdAt"] as? Timestamp)?.dateValue()
                        let claimedAt = (data["claimedAt"] as? Timestamp)?.dateValue()

                        return FamilyInvite(
                            id: document.documentID,
                            code: data["code"] as? String ?? "",
                            email: data["inviteEmail"] as? String ?? "",
                            inviteLink: data["inviteLink"] as? String ?? "",
                            createdAt: createdAt,
                            claimedAt: claimedAt
                        )
                    }
                    .sorted { lhs, rhs in
                        (lhs.createdAt ?? .distantPast) > (rhs.createdAt ?? .distantPast)
                    }
                }
            }
    }

    func stopListening() {
        membersListener?.remove()
        membersListener = nil
        invitesListener?.remove()
        invitesListener = nil
        currentFamilyId = nil
    }

    func createInvite(
        familyId: String,
        createdByUserId: String
    ) async -> Bool {
        let normalizedEmail = inviteEmail.trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
        guard normalizedEmail.contains("@") else {
            errorMessage = "Please enter a valid email."
            return false
        }
        guard !familyId.isEmpty else {
            errorMessage = "Family is not configured yet."
            return false
        }

        isCreatingInvite = true
        errorMessage = nil
        defer { isCreatingInvite = false }

        do {
            let inviteCode = try await generateUniqueInviteCode()
            let inviteLink = "familydinnerplanner://signup?inviteCode=\(inviteCode)"

            try await db.collection("invites").addDocument(data: [
                "code": inviteCode,
                "familyId": familyId,
                "inviteEmail": normalizedEmail,
                "inviteEmailLowercase": normalizedEmail,
                "createdBy": createdByUserId,
                "createdAt": FieldValue.serverTimestamp(),
                "status": "active",
                "claimedBy": "",
                "inviteLink": inviteLink
            ])

            latestInviteCode = inviteCode
            latestInviteLink = inviteLink
            inviteEmail = ""
            return true
        } catch {
            errorMessage = error.localizedDescription
            return false
        }
    }

    private func resolveName(from value: String?, email: String, fallback: String) -> String {
        let name = value?.trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
        if !name.isEmpty {
            return name
        }

        if let localPart = email.split(separator: "@").first, !localPart.isEmpty {
            return String(localPart)
        }

        return fallback
    }

    private func generateUniqueInviteCode() async throws -> String {
        var attempts = 0

        while attempts < 8 {
            attempts += 1
            let code = Self.randomCode(length: 8)
            let existing = try await db
                .collection("invites")
                .whereField("code", isEqualTo: code)
                .limit(to: 1)
                .getDocuments()

            if existing.documents.isEmpty {
                return code
            }
        }

        throw NSError(
            domain: "FamilyManagement",
            code: 1,
            userInfo: [NSLocalizedDescriptionKey: "Unable to generate invite code. Try again."]
        )
    }

    private static func randomCode(length: Int) -> String {
        let charset = Array("ABCDEFGHJKLMNPQRSTUVWXYZ23456789")
        return String((0..<length).compactMap { _ in charset.randomElement() })
    }
}
