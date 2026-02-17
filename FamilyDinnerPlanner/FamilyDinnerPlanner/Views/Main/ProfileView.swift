import SwiftUI
import Observation

struct ProfileView: View {
    @Bindable var authService: AuthService

    var body: some View {
        Form {
            Section {
                LabeledContent("Email", value: authService.currentUser?.email ?? "Unknown")
                LabeledContent("User ID", value: authService.currentUser?.uid ?? "Not signed in")
                LabeledContent(
                    "Family ID",
                    value: authService.currentFamilyId ?? "Not set"
                )
            } header: {
                Label("Account", systemImage: "person.crop.circle.fill")
            }

            Section {
                LabeledContent(
                    "Role",
                    value: authService.userRole?.rawValue.capitalized ?? "Unknown"
                )
                Text("Current profile name: \(authService.currentDisplayName ?? "Family Member")")
                    .foregroundStyle(.secondary)
            } header: {
                Label("Membership", systemImage: "house.fill")
            }

            Section {
                Button(role: .destructive) {
                    _ = authService.signOut()
                } label: {
                    Label("Sign Out", systemImage: "rectangle.portrait.and.arrow.right")
                }
            }
        }
        .navigationTitle("Profile")
        .toolbar {
            ToolbarItem(placement: .topBarTrailing) {
                Image(systemName: "person.circle.fill")
                    .foregroundStyle(AppTheme.accent)
            }
        }
    }
}
