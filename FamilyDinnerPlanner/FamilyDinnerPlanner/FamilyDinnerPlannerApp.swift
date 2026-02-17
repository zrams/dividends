import SwiftUI
import FirebaseCore
import FirebaseAuth
import FirebaseFirestore
import FirebaseMessaging

@main
struct FamilyDinnerPlannerApp: App {
    @UIApplicationDelegateAdaptor(AppDelegate.self) private var appDelegate

    @State private var authService: AuthService
    @State private var firestoreService: FirestoreService

    init() {
        FirebaseApp.configure()

        // Eagerly initialize Firebase modules used by the app.
        _ = Auth.auth()
        let firestore = Firestore.firestore()
        let firestoreSettings = FirestoreSettings()
        firestoreSettings.isPersistenceEnabled = true
        firestoreSettings.cacheSizeBytes = FirestoreCacheSizeUnlimited
        firestore.settings = firestoreSettings
        _ = Messaging.messaging()

        _authService = State(initialValue: AuthService())
        _firestoreService = State(initialValue: FirestoreService())
    }

    var body: some Scene {
        WindowGroup {
            LaunchRouterView(
                authService: authService,
                firestoreService: firestoreService
            )
            .tint(AppTheme.accent)
        }
    }
}
