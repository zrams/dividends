import SwiftUI
import Observation

struct MemberSubmissionView: View {
    @Bindable var authService: AuthService
    @State private var viewModel = MemberSubmissionViewModel()
    @State private var showConfirmationAlert = false
    @State private var weekStartDate = MemberSubmissionView.nextMonday(after: .now)

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

    private var formattedWeekStart: String {
        let formatter = DateFormatter()
        formatter.dateStyle = .medium
        formatter.timeStyle = .none
        return formatter.string(from: weekStartDate)
    }

    var body: some View {
        Group {
            if authService.userRole != .member {
                ContentUnavailableView(
                    "Members Only",
                    systemImage: "lock.circle",
                    description: Text("Only members can submit weekly choices.")
                )
            } else if viewModel.isLoadingDinners && viewModel.dinners.isEmpty {
                ProgressView("Loading dinners...")
            } else {
                List {
                    Section("Upcoming Week") {
                        Label("Week starts on \(formattedWeekStart) (next Monday).", systemImage: "calendar.badge.clock")
                            .font(.subheadline)
                            .foregroundStyle(.secondary)

                        if viewModel.isCheckingExistingSubmission {
                            ProgressView("Checking existing submission...")
                        } else if viewModel.hasExistingSubmissionForWeek {
                            Label("You already submitted for this week.", systemImage: "checkmark.seal.fill")
                                .foregroundStyle(.green)
                        } else {
                            Text("Choose 1 or 2 dinners. Selected: \(viewModel.selectedDinnerIDs.count)/2.")
                                .font(.subheadline)
                                .foregroundStyle(.secondary)
                        }
                    }

                    Section("Dinner Ideas") {
                        if viewModel.dinners.isEmpty {
                            ContentUnavailableView(
                                "No Dinners Available",
                                systemImage: "fork.knife",
                                description: Text("Ask an admin to add dinners, or check your connection for cached data.")
                            )
                        } else {
                            ForEach(viewModel.dinners) { dinner in
                                MemberDinnerRow(
                                    dinner: dinner,
                                    isSelected: viewModel.selectedDinnerIDs.contains(dinner.id)
                                )
                                .contentShape(Rectangle())
                                .onTapGesture {
                                    guard !viewModel.hasExistingSubmissionForWeek else { return }
                                    guard !viewModel.isSubmitting else { return }
                                    viewModel.toggleSelection(for: dinner.id)
                                }
                            }
                        }
                    }

                    Section {
                        Button {
                            submitSelection()
                        } label: {
                            if viewModel.isSubmitting {
                                ProgressView()
                                    .frame(maxWidth: .infinity)
                            } else {
                                Text("Submit Choices")
                                    .frame(maxWidth: .infinity)
                            }
                        }
                        .disabled(!viewModel.canSubmit || authService.currentUser == nil)
                    } footer: {
                        Text("Choices are saved in Firestore 'submissions' for the upcoming week.")
                    }
                }
                .listStyle(.insetGrouped)
            }
        }
        .navigationTitle("Member Submission")
        .toolbar {
            ToolbarItem(placement: .topBarTrailing) {
                Image(systemName: "fork.knife.circle.fill")
                    .foregroundStyle(AppTheme.accent)
            }
        }
        .task {
            guard let userId = authService.currentUser?.uid else { return }
            await viewModel.loadInitialState(userId: userId, weekStartDate: weekStartDate)
        }
        .refreshable {
            guard let userId = authService.currentUser?.uid else { return }
            await viewModel.loadInitialState(userId: userId, weekStartDate: weekStartDate)
        }
        .alert(
            "Unable to Submit",
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
        .alert("Submission Received", isPresented: $showConfirmationAlert) {
            Button("OK", role: .cancel) {}
        } message: {
            Text("Your weekly dinner choices were submitted.")
        }
    }

    private func submitSelection() {
        guard let userId = authService.currentUser?.uid else { return }
        guard let familyId = authService.currentFamilyId, !familyId.isEmpty else {
            viewModel.errorMessage = "Family setup is incomplete. Please sign out and back in."
            return
        }

        Task {
            let success = await viewModel.submit(
                userId: userId,
                weekStartDate: weekStartDate,
                familyId: familyId
            )
            if success {
                showConfirmationAlert = true
            }
        }
    }
}

private struct MemberDinnerRow: View {
    let dinner: DinnerIdea
    let isSelected: Bool

    var body: some View {
        HStack(spacing: 12) {
            VStack(alignment: .leading, spacing: 4) {
                Text(dinner.name)
                    .font(.headline)

                if let description = dinner.description, !description.isEmpty {
                    Text(description)
                        .font(.subheadline)
                        .foregroundStyle(.secondary)
                }
            }

            Spacer()

            Image(systemName: isSelected ? "checkmark.circle.fill" : "circle")
                .foregroundStyle(isSelected ? .green : .secondary)
        }
        .padding(.vertical, 4)
    }
}

private extension MemberSubmissionView {
    static func nextMonday(after date: Date) -> Date {
        let calendar = Calendar.current
        let startOfToday = calendar.startOfDay(for: date)
        let weekday = calendar.component(.weekday, from: startOfToday)
        let daysUntilNextMonday = weekday == 2 ? 7 : ((2 - weekday + 7) % 7)
        let candidateDate = calendar.date(byAdding: .day, value: daysUntilNextMonday, to: startOfToday) ?? startOfToday

        var components = calendar.dateComponents([.yearForWeekOfYear, .weekOfYear], from: candidateDate)
        components.weekday = 2
        return calendar.date(from: components) ?? candidateDate
    }
}
