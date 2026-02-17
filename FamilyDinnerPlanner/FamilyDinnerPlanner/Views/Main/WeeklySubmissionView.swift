import SwiftUI
import Observation

struct WeeklySubmissionView: View {
    @Bindable var authService: AuthService
    @Bindable var firestoreService: FirestoreService

    @State private var selectedDinnerIDs: Set<String> = []
    @State private var isSubmitting = false
    @State private var showSuccessAlert = false

    private var showErrorAlert: Binding<Bool> {
        Binding(
            get: { firestoreService.firestoreErrorMessage != nil },
            set: { newValue in
                if !newValue {
                    firestoreService.firestoreErrorMessage = nil
                }
            }
        )
    }

    private var weekLabel: String {
        let formatter = DateFormatter()
        formatter.dateStyle = .medium
        formatter.timeStyle = .none
        return formatter.string(from: .now)
    }

    var body: some View {
        Form {
            Section("Week of \(weekLabel)") {
                if firestoreService.dinnerIdeas.isEmpty {
                    Text("No dinner ideas available yet.")
                        .foregroundStyle(.secondary)
                } else {
                    ForEach(firestoreService.dinnerIdeas) { dinnerIdea in
                        Toggle(isOn: binding(for: dinnerIdea.id)) {
                            VStack(alignment: .leading, spacing: 4) {
                                Text(dinnerIdea.name)
                                if let description = dinnerIdea.description, !description.isEmpty {
                                    Text(description)
                                        .font(.caption)
                                        .foregroundStyle(.secondary)
                                }
                            }
                        }
                    }
                }
            }

            Section {
                Button {
                    submitWeeklyChoices()
                } label: {
                    if isSubmitting {
                        ProgressView()
                            .frame(maxWidth: .infinity)
                    } else {
                        Text("Submit Choices")
                            .frame(maxWidth: .infinity)
                    }
                }
                .disabled(
                    isSubmitting ||
                    selectedDinnerIDs.isEmpty ||
                    authService.currentUser == nil
                )
            } footer: {
                Text("Selected choices are saved in Firestore as a WeeklySubmission document.")
            }
        }
        .navigationTitle("Weekly Submission")
        .task {
            if firestoreService.dinnerIdeas.isEmpty {
                await firestoreService.fetchDinnerIdeas()
            }
        }
        .alert("Submitted", isPresented: $showSuccessAlert) {
            Button("OK", role: .cancel) { }
        } message: {
            Text("Your dinner choices were submitted successfully.")
        }
        .alert(
            "Submission Failed",
            isPresented: showErrorAlert,
            actions: {
                Button("OK") {
                    firestoreService.firestoreErrorMessage = nil
                }
            },
            message: {
                Text(firestoreService.firestoreErrorMessage ?? "Unknown error.")
            }
        )
    }

    private func binding(for dinnerID: String) -> Binding<Bool> {
        Binding(
            get: { selectedDinnerIDs.contains(dinnerID) },
            set: { isSelected in
                if isSelected {
                    selectedDinnerIDs.insert(dinnerID)
                } else {
                    selectedDinnerIDs.remove(dinnerID)
                }
            }
        )
    }

    private func submitWeeklyChoices() {
        guard let userId = authService.currentUser?.uid else { return }
        isSubmitting = true

        Task {
            defer { isSubmitting = false }

            let success = await firestoreService.submitWeeklySubmission(
                userId: userId,
                choices: Array(selectedDinnerIDs).sorted()
            )

            if success {
                showSuccessAlert = true
                selectedDinnerIDs = []
            }
        }
    }
}
