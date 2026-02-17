import SwiftUI
import Observation

struct MainTabView: View {
    enum Tab {
        case ideas
        case submit
        case admin
        case profile
    }

    @Bindable var authService: AuthService
    @Bindable var firestoreService: FirestoreService
    @State private var selectedTab: Tab = .ideas

    private var showRoleErrorAlert: Binding<Bool> {
        Binding(
            get: { authService.roleErrorMessage != nil },
            set: { isPresented in
                if !isPresented {
                    authService.roleErrorMessage = nil
                }
            }
        )
    }

    var body: some View {
        Group {
            if authService.isLoadingUserRole {
                ProgressView("Loading account role...")
            } else {
                TabView(selection: $selectedTab) {
                    NavigationStack {
                        HomeView(firestoreService: firestoreService)
                    }
                    .tabItem {
                        Label("Ideas", systemImage: "list.bullet")
                    }
                    .tag(Tab.ideas)

                    if authService.userRole == .member {
                        NavigationStack {
                            MemberSubmissionView(authService: authService)
                        }
                        .tabItem {
                            Label("Weekly Pick", systemImage: "calendar.badge.plus")
                        }
                        .tag(Tab.submit)
                    }

                    if authService.isAdmin {
                        NavigationStack {
                            AdminDashboardView(authService: authService)
                        }
                        .tabItem {
                            Label("Admin", systemImage: "person.3.sequence.fill")
                        }
                        .tag(Tab.admin)
                    }

                    NavigationStack {
                        ProfileView(authService: authService)
                    }
                    .tabItem {
                        Label("Profile", systemImage: "person.circle")
                    }
                    .tag(Tab.profile)
                }
            }
        }
        .task {
            await NotificationPermissionService.requestIfNeeded()
        }
        .onChange(of: authService.userRole) { _, newRole in
            if newRole != .admin && selectedTab == .admin {
                selectedTab = .ideas
            }
            if newRole != .member && selectedTab == .submit {
                selectedTab = .ideas
            }
        }
        .alert(
            "Unable to Load Role",
            isPresented: showRoleErrorAlert,
            actions: {
                Button("OK", role: .cancel) {
                    authService.roleErrorMessage = nil
                }
            },
            message: {
                Text(authService.roleErrorMessage ?? "Unknown error.")
            }
        )
    }
}
