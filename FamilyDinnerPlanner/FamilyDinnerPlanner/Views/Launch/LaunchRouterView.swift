import SwiftUI
import Observation

struct LaunchRouterView: View {
    @Bindable var authService: AuthService
    let firestoreService: FirestoreService

    var body: some View {
        Group {
            switch authService.sessionState {
            case .loading:
                LaunchLoadingView()
            case .signedOut:
                NavigationStack {
                    LoginView(authService: authService)
                }
            case .signedIn:
                MainTabView(
                    authService: authService,
                    firestoreService: firestoreService
                )
            }
        }
        .animation(.snappy, value: authService.isAuthenticated)
    }
}

private struct LaunchLoadingView: View {
    var body: some View {
        ZStack {
            LinearGradient(
                colors: [.orange.opacity(0.25), .red.opacity(0.2), .yellow.opacity(0.15)],
                startPoint: .topLeading,
                endPoint: .bottomTrailing
            )
            .ignoresSafeArea()

            VStack(spacing: 16) {
                Image(systemName: "fork.knife.circle.fill")
                    .font(.system(size: 68))
                    .foregroundStyle(.orange)
                    .symbolRenderingMode(.hierarchical)

                Text("FamilyDinnerPlanner")
                    .font(.largeTitle.bold())

                ProgressView("Loading...")
                    .tint(.orange)
            }
            .padding()
        }
    }
}
