import SwiftUI
import Observation

struct FamilyManagementView: View {
    @Bindable var authService: AuthService
    @State private var viewModel = FamilyManagementViewModel()
    @State private var showInviteCreatedAlert = false

    private var showErrorAlert: Binding<Bool> {
        Binding(
            get: { viewModel.errorMessage != nil },
            set: { isPresented in
                if !isPresented {
                    viewModel.errorMessage = nil
                }
            }
        )
    }

    var body: some View {
        Group {
            if !authService.isAdmin {
                ContentUnavailableView(
                    "Admin Only",
                    systemImage: "lock.shield",
                    description: Text("Family management is available to admins only.")
                )
            } else if authService.currentFamilyId == nil {
                ContentUnavailableView(
                    "Family Setup Pending",
                    systemImage: "house",
                    description: Text("Sign out and back in if family configuration does not appear.")
                )
            } else if viewModel.isLoadingMembers && viewModel.members.isEmpty {
                ProgressView("Loading family members...")
            } else {
                List {
                    Section {
                        HStack(alignment: .center, spacing: 10) {
                            Image(systemName: "envelope.badge.fill")
                                .foregroundStyle(AppTheme.accent)
                            TextField("Invite member email", text: $viewModel.inviteEmail)
                                .textInputAutocapitalization(.never)
                                .keyboardType(.emailAddress)
                                .autocorrectionDisabled(true)
                        }

                        Button {
                            createInvite()
                        } label: {
                            if viewModel.isCreatingInvite {
                                ProgressView()
                                    .frame(maxWidth: .infinity)
                            } else {
                                Label("Create Invite", systemImage: "paperplane.fill")
                                    .frame(maxWidth: .infinity)
                            }
                        }
                        .disabled(
                            viewModel.isCreatingInvite ||
                            viewModel.inviteEmail.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty
                        )
                    } header: {
                        Label("Invite Family Member", systemImage: "person.badge.plus")
                    } footer: {
                        Text("Invite links are saved in Firestore 'invites'. Share the code/link with your family member.")
                    }

                    if let latestCode = viewModel.latestInviteCode,
                       let latestLink = viewModel.latestInviteLink {
                        Section {
                            LabeledContent("Code", value: latestCode)
                            ShareLink(item: latestLink) {
                                Label("Share Invite Link", systemImage: "square.and.arrow.up")
                            }
                        } header: {
                            Label("Latest Invite", systemImage: "link.badge.plus")
                        }
                    }

                    Section {
                        if viewModel.members.isEmpty {
                            ContentUnavailableView(
                                "No Family Members Yet",
                                systemImage: "person.2.slash",
                                description: Text("Invite someone to join your family group.")
                            )
                        } else {
                            ForEach(viewModel.members) { member in
                                VStack(alignment: .leading, spacing: 4) {
                                    Label(member.name, systemImage: "person.fill")
                                        .font(.headline)
                                    Text(member.email)
                                        .font(.subheadline)
                                        .foregroundStyle(.secondary)
                                }
                                .padding(.vertical, 4)
                            }
                        }
                    } header: {
                        Label("Family Members (\(viewModel.members.count))", systemImage: "person.3.fill")
                    }

                    if !viewModel.invites.isEmpty {
                        Section {
                            ForEach(viewModel.invites.prefix(10)) { invite in
                                VStack(alignment: .leading, spacing: 4) {
                                    Text(invite.email)
                                        .font(.headline)
                                    Text("Code: \(invite.code)")
                                        .font(.caption)
                                        .foregroundStyle(.secondary)
                                    Text(invite.isClaimed ? "Claimed" : "Pending")
                                        .font(.caption2.weight(.semibold))
                                        .foregroundStyle(invite.isClaimed ? .green : .orange)
                                }
                                .padding(.vertical, 2)
                            }
                        } header: {
                            Label("Recent Invites", systemImage: "envelope.open.fill")
                        }
                    }
                }
                .listStyle(.insetGrouped)
            }
        }
        .navigationTitle("Family Management")
        .toolbar {
            ToolbarItem(placement: .topBarTrailing) {
                Button {
                    refreshFamilyData()
                } label: {
                    Image(systemName: "arrow.clockwise")
                }
            }
        }
        .task {
            refreshFamilyData()
        }
        .onChange(of: authService.currentFamilyId) { _, _ in
            refreshFamilyData()
        }
        .alert(
            "Unable to Update Family",
            isPresented: showErrorAlert,
            actions: {
                Button("OK", role: .cancel) {
                    viewModel.errorMessage = nil
                }
            },
            message: {
                Text(viewModel.errorMessage ?? "Unknown error.")
            }
        )
        .alert("Invite Created", isPresented: $showInviteCreatedAlert) {
            Button("OK", role: .cancel) {}
        } message: {
            Text("Invite code and link were saved. Share them with your family member.")
        }
    }

    private func refreshFamilyData() {
        guard let familyId = authService.currentFamilyId else {
            viewModel.stopListening()
            return
        }
        viewModel.startListening(familyId: familyId)
    }

    private func createInvite() {
        guard let familyId = authService.currentFamilyId,
              let adminID = authService.currentUser?.uid else {
            return
        }

        Task {
            let success = await viewModel.createInvite(
                familyId: familyId,
                createdByUserId: adminID
            )
            if success {
                showInviteCreatedAlert = true
            }
        }
    }
}
