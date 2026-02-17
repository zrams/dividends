import SwiftUI
import Observation

struct MainTabView: View {
    enum Tab {
        case ideas
        case submit
        case history
        case admin
        case family
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

                        NavigationStack {
                            MemberHistoryView(authService: authService)
                        }
                        .tabItem {
                            Label("History", systemImage: "clock.arrow.circlepath")
                        }
                        .tag(Tab.history)
                    }

                    if authService.isAdmin {
                        NavigationStack {
                            AdminDashboardView(authService: authService)
                        }
                        .tabItem {
                            Label("Admin", systemImage: "slider.horizontal.3")
                        }
                        .tag(Tab.admin)

                        NavigationStack {
                            FamilyManagementView(authService: authService)
                        }
                        .tabItem {
                            Label("Family", systemImage: "person.2.fill")
                        }
                        .tag(Tab.family)
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
        .onOpenURL { url in
            handleDeepLink(url)
        }
        .onChange(of: authService.userRole) { _, newRole in
            if newRole != .admin && selectedTab == .admin {
                selectedTab = .ideas
            }
            if newRole != .admin && selectedTab == .family {
                selectedTab = .ideas
            }
            if newRole != .member && selectedTab == .submit {
                selectedTab = .ideas
            }
            if newRole != .member && selectedTab == .history {
                selectedTab = .ideas
            }
        }
        .tint(AppTheme.accent)
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

    private func handleDeepLink(_ url: URL) {
        guard url.scheme?.lowercased() == "familydinnerplanner" else { return }

        let destination = "\(url.host ?? "")\(url.path)".lowercased()
        if destination.contains("member-submission") && authService.userRole == .member {
            selectedTab = .submit
        }
    }
}
