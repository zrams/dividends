import Foundation
import FirebaseAuth
import FirebaseFirestore
import FirebaseMessaging
import Observation

@MainActor
@Observable
final class AuthService {
    enum InviteClaimError: LocalizedError {
        case invalidCode
        case emailMismatch
        case missingFamilyId
        case alreadyClaimed

        var errorDescription: String? {
            switch self {
            case .invalidCode:
                return "That invite code is not valid."
            case .emailMismatch:
                return "This invite was created for a different email address."
            case .missingFamilyId:
                return "Invite is missing a family. Ask your admin to create a new invite."
            case .alreadyClaimed:
                return "This invite code has already been claimed."
            }
        }
    }

    private struct InviteContext {
        let inviteDocumentID: String
        let familyId: String
    }

    enum SessionState {
        case loading
        case signedIn(User)
        case signedOut
    }

    var sessionState: SessionState = .loading
    var authErrorMessage: String?
    var userRole: UserRole?
    var isLoadingUserRole = false
    var roleErrorMessage: String?
    var currentFamilyId: String?
    var currentDisplayName: String?

    private var authStateListener: AuthStateDidChangeListenerHandle?
    private var userRoleListener: ListenerRegistration?
    private let db = Firestore.firestore()

    init() {
        startAuthStateListener()
    }

    var currentUser: User? {
        guard case let .signedIn(user) = sessionState else {
            return nil
        }
        return user
    }

    var isAuthenticated: Bool {
        currentUser != nil
    }

    var isAdmin: Bool {
        userRole == .admin
    }

    func signIn(email: String, password: String) async -> Bool {
        authErrorMessage = nil

        do {
            _ = try await Auth.auth().signIn(withEmail: email, password: password)
            return true
        } catch {
            authErrorMessage = error.localizedDescription
            return false
        }
    }

    func signUp(
        email: String,
        password: String,
        name: String,
        inviteCode: String?
    ) async -> Bool {
        authErrorMessage = nil

        let normalizedEmail = email.trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
        let normalizedName = name.trimmingCharacters(in: .whitespacesAndNewlines)
        let normalizedInviteCode = inviteCode?.trimmingCharacters(in: .whitespacesAndNewlines) ?? ""

        var createdUser: User?

        do {
            let inviteContext = try await resolveInviteContextIfNeeded(
                inviteCode: normalizedInviteCode,
                emailLowercase: normalizedEmail
            )

            let result = try await Auth.auth().createUser(withEmail: email, password: password)
            createdUser = result.user

            if let inviteContext {
                try await claimInvite(
                    inviteContext: inviteContext,
                    userID: result.user.uid,
                    emailLowercase: normalizedEmail
                )
            }

            let assignedFamilyID = inviteContext?.familyId ?? result.user.uid

            var profileData: [String: Any] = [
                "role": UserRole.member.rawValue,
                "email": normalizedEmail,
                "familyId": assignedFamilyID,
                "updatedAt": FieldValue.serverTimestamp(),
                "createdAt": FieldValue.serverTimestamp()
            ]

            if !normalizedName.isEmpty {
                profileData["name"] = normalizedName
            }

            try await db
                .collection("users")
                .document(result.user.uid)
                .setData(profileData, merge: true)

            return true
        } catch {
            if let createdUser {
                try? await createdUser.delete()
            }
            authErrorMessage = error.localizedDescription
            return false
        }
    }

    func signOut() -> Bool {
        authErrorMessage = nil

        do {
            try Auth.auth().signOut()
            return true
        } catch {
            authErrorMessage = error.localizedDescription
            return false
        }
    }

    private func startAuthStateListener() {
        authStateListener = Auth.auth().addStateDidChangeListener { [weak self] _, user in
            Task { @MainActor [weak self] in
                guard let self else { return }
                if let user {
                    self.sessionState = .signedIn(user)
                    self.startUserRoleListener(for: user)
                    Task {
                        await self.syncCurrentFCMToken(for: user.uid)
                    }
                } else {
                    self.sessionState = .signedOut
                    self.stopUserRoleListener()
                    self.userRole = nil
                    self.isLoadingUserRole = false
                    self.roleErrorMessage = nil
                    self.currentFamilyId = nil
                    self.currentDisplayName = nil
                }
            }
        }
    }

    private func startUserRoleListener(for user: User) {
        stopUserRoleListener()
        isLoadingUserRole = true
        roleErrorMessage = nil

        userRoleListener = db.collection("users").document(user.uid).addSnapshotListener { [weak self] snapshot, error in
            Task { @MainActor [weak self] in
                guard let self else { return }

                if let error {
                    self.roleErrorMessage = error.localizedDescription
                    self.userRole = .member
                    self.currentFamilyId = user.uid
                    self.currentDisplayName = self.fallbackDisplayName(
                        from: nil,
                        email: user.email
                    )
                    self.isLoadingUserRole = false
                    return
                }

                let data = snapshot?.data() ?? [:]
                let roleRawValue = (data["role"] as? String)?.lowercased()
                self.userRole = UserRole(rawValue: roleRawValue ?? "") ?? .member

                let familyID = self.nonEmptyString(from: data["familyId"] as? String) ?? user.uid
                self.currentFamilyId = familyID

                if self.nonEmptyString(from: data["familyId"] as? String) == nil {
                    Task {
                        await self.ensureFamilyIdExists(for: user.uid, familyId: user.uid)
                    }
                }

                self.currentDisplayName = self.fallbackDisplayName(
                    from: data["name"] as? String,
                    email: (data["email"] as? String) ?? user.email
                )
                self.isLoadingUserRole = false
            }
        }
    }

    private func stopUserRoleListener() {
        userRoleListener?.remove()
        userRoleListener = nil
    }

    private func syncCurrentFCMToken(for uid: String) async {
        do {
            let fcmToken = try await fetchCurrentFCMToken()
            guard !fcmToken.isEmpty else { return }

            try await db.collection("users").document(uid).setData(
                ["fcmToken": fcmToken],
                merge: true
            )
        } catch {
            print("FCM sync failed: \(error.localizedDescription)")
        }
    }

    private func fetchCurrentFCMToken() async throws -> String {
        try await withCheckedThrowingContinuation { continuation in
            Messaging.messaging().token { token, error in
                if let error {
                    continuation.resume(throwing: error)
                    return
                }

                continuation.resume(returning: token ?? "")
            }
        }
    }

    private func resolveInviteContextIfNeeded(
        inviteCode: String,
        emailLowercase: String
    ) async throws -> InviteContext? {
        guard !inviteCode.isEmpty else { return nil }

        let normalizedCode = inviteCode.uppercased()
        let snapshot = try await db
            .collection("invites")
            .whereField("code", isEqualTo: normalizedCode)
            .limit(to: 1)
            .getDocuments()

        guard let inviteDocument = snapshot.documents.first else {
            throw InviteClaimError.invalidCode
        }

        let data = inviteDocument.data()
        let status = (data["status"] as? String)?.lowercased() ?? "active"
        if status != "active" {
            throw InviteClaimError.alreadyClaimed
        }
        if let claimedBy = nonEmptyString(from: data["claimedBy"] as? String), !claimedBy.isEmpty {
            throw InviteClaimError.alreadyClaimed
        }

        let inviteEmail = nonEmptyString(from: (data["inviteEmailLowercase"] as? String) ?? (data["emailLowercase"] as? String))
        if let inviteEmail, inviteEmail != emailLowercase {
            throw InviteClaimError.emailMismatch
        }

        guard let familyID = nonEmptyString(from: data["familyId"] as? String) else {
            throw InviteClaimError.missingFamilyId
        }

        return InviteContext(
            inviteDocumentID: inviteDocument.documentID,
            familyId: familyID
        )
    }

    private func claimInvite(
        inviteContext: InviteContext,
        userID: String,
        emailLowercase: String
    ) async throws {
        let inviteReference = db.collection("invites").document(inviteContext.inviteDocumentID)
        let latestSnapshot = try await inviteReference.getDocument()
        let latestData = latestSnapshot.data() ?? [:]
        let status = (latestData["status"] as? String)?.lowercased() ?? "active"
        if status != "active" {
            throw InviteClaimError.alreadyClaimed
        }

        if let claimedBy = nonEmptyString(from: latestData["claimedBy"] as? String), !claimedBy.isEmpty {
            throw InviteClaimError.alreadyClaimed
        }

        try await inviteReference.setData([
            "claimedBy": userID,
            "claimedEmailLowercase": emailLowercase,
            "claimedAt": FieldValue.serverTimestamp(),
            "status": "claimed"
        ], merge: true)
    }

    private func ensureFamilyIdExists(for userID: String, familyId: String) async {
        do {
            try await db.collection("users").document(userID).setData(
                ["familyId": familyId],
                merge: true
            )
        } catch {
            print("Failed to write default familyId: \(error.localizedDescription)")
        }
    }

    private func nonEmptyString(from value: String?) -> String? {
        let normalized = value?.trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
        return normalized.isEmpty ? nil : normalized
    }

    private func fallbackDisplayName(from value: String?, email: String?) -> String {
        if let value = nonEmptyString(from: value) {
            return value
        }

        if let email = nonEmptyString(from: email) {
            if let localPart = email.split(separator: "@").first {
                return String(localPart)
            }
            return email
        }

        return "Family Member"
    }
}
