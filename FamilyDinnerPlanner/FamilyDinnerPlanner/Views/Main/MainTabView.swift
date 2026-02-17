import SwiftUI
import Observation

struct MainTabView: View {
    enum Tab {
        case ideas
        case submit
        case profile
    }

    @Bindable var authService: AuthService
    @Bindable var firestoreService: FirestoreService
    @State private var selectedTab: Tab = .ideas

    var body: some View {
        TabView(selection: $selectedTab) {
            NavigationStack {
                HomeView(firestoreService: firestoreService)
            }
            .tabItem {
                Label("Ideas", systemImage: "list.bullet")
            }
            .tag(Tab.ideas)

            NavigationStack {
                WeeklySubmissionView(
                    authService: authService,
                    firestoreService: firestoreService
                )
            }
            .tabItem {
                Label("Weekly Pick", systemImage: "calendar.badge.plus")
            }
            .tag(Tab.submit)

            NavigationStack {
                ProfileView(authService: authService)
            }
            .tabItem {
                Label("Profile", systemImage: "person.circle")
            }
            .tag(Tab.profile)
        }
        .task {
            await NotificationPermissionService.requestIfNeeded()
        }
    }
}
