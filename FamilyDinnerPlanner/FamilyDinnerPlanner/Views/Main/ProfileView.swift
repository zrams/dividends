import SwiftUI
import Observation

struct ProfileView: View {
    @Bindable var authService: AuthService

    var body: some View {
        Form {
            Section("Account") {
                LabeledContent("Email", value: authService.currentUser?.email ?? "Unknown")
                LabeledContent("User ID", value: authService.currentUser?.uid ?? "Not signed in")
            }

            Section("Role") {
                Text("Default role: \(UserRole.member.rawValue)")
                    .foregroundStyle(.secondary)
            }

            Section {
                Button("Sign Out", role: .destructive) {
                    _ = authService.signOut()
                }
            }
        }
        .navigationTitle("Profile")
    }
}
