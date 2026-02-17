import Foundation
import FirebaseAuth
import FirebaseFirestore
import FirebaseMessaging
import Observation

@MainActor
@Observable
final class AuthService {
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

    func signUp(email: String, password: String) async -> Bool {
        authErrorMessage = nil

        do {
            let result = try await Auth.auth().createUser(withEmail: email, password: password)

            // Create the profile document on first sign-up with a default member role.
            try? await db
                .collection("users")
                .document(result.user.uid)
                .setData(["role": UserRole.member.rawValue], merge: true)

            return true
        } catch {
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
                    self.startUserRoleListener(for: user.uid)
                    Task {
                        await self.syncCurrentFCMToken(for: user.uid)
                    }
                } else {
                    self.sessionState = .signedOut
                    self.stopUserRoleListener()
                    self.userRole = nil
                    self.isLoadingUserRole = false
                    self.roleErrorMessage = nil
                }
            }
        }
    }

    private func startUserRoleListener(for uid: String) {
        stopUserRoleListener()
        isLoadingUserRole = true
        roleErrorMessage = nil

        userRoleListener = db.collection("users").document(uid).addSnapshotListener { [weak self] snapshot, error in
            Task { @MainActor [weak self] in
                guard let self else { return }

                if let error {
                    self.roleErrorMessage = error.localizedDescription
                    self.userRole = .member
                    self.isLoadingUserRole = false
                    return
                }

                let roleRawValue = (snapshot?.data()?["role"] as? String)?.lowercased()
                self.userRole = UserRole(rawValue: roleRawValue ?? "") ?? .member
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
}
