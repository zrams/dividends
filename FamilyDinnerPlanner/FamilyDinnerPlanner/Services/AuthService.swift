import Foundation
import FirebaseAuth
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

    private var authStateListener: AuthStateDidChangeListenerHandle?

    init() {
        startAuthStateListener()
    }

    deinit {
        if let authStateListener {
            Auth.auth().removeStateDidChangeListener(authStateListener)
        }
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
            _ = try await Auth.auth().createUser(withEmail: email, password: password)
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
                } else {
                    self.sessionState = .signedOut
                }
            }
        }
    }
}
