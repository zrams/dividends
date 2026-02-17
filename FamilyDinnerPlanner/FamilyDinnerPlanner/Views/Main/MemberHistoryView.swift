import SwiftUI
import Observation

struct MemberHistoryView: View {
    @Bindable var authService: AuthService
    @State private var viewModel = MemberHistoryViewModel()

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
            if authService.userRole != .member {
                ContentUnavailableView(
                    "Members Only",
                    systemImage: "clock.badge.questionmark",
                    description: Text("Submission history is available to members only.")
                )
            } else if viewModel.isLoading && viewModel.historyItems.isEmpty {
                ProgressView("Loading submission history...")
            } else if viewModel.historyItems.isEmpty {
                ContentUnavailableView(
                    "No Submission History",
                    systemImage: "clock.arrow.circlepath",
                    description: Text("Your previous weekly choices will appear here.")
                )
            } else {
                List {
                    ForEach(viewModel.historyItems) { item in
                        Section {
                            if item.choiceNames.isEmpty {
                                Text("No choices recorded for this week.")
                                    .foregroundStyle(.secondary)
                            } else {
                                ForEach(item.choiceNames, id: \.self) { choice in
                                    Label(choice, systemImage: "fork.knife")
                                }
                            }
                        } header: {
                            Label(
                                "Week of \(formattedDate(item.weekStart))",
                                systemImage: "calendar"
                            )
                        }
                    }
                }
                .listStyle(.insetGrouped)
            }
        }
        .navigationTitle("Past Weeks")
        .task {
            guard let userID = authService.currentUser?.uid else { return }
            viewModel.startListening(userId: userID)
        }
        .onChange(of: authService.currentUser?.uid) { _, newUserID in
            viewModel.stopListening()
            if let newUserID {
                viewModel.startListening(userId: newUserID)
            }
        }
        .onDisappear {
            viewModel.stopListening()
        }
        .alert(
            "Unable to Load History",
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
    }

    private func formattedDate(_ date: Date) -> String {
        let formatter = DateFormatter()
        formatter.dateStyle = .medium
        formatter.timeStyle = .none
        return formatter.string(from: date)
    }
}
