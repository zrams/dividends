import Foundation
import FirebaseFirestore
import Observation

@MainActor
@Observable
final class MemberSubmissionViewModel {
    private let db = Firestore.firestore()

    var dinners: [DinnerIdea] = []
    var selectedDinnerIDs: Set<String> = []
    var isLoadingDinners = false
    var isCheckingExistingSubmission = false
    var isSubmitting = false
    var hasExistingSubmissionForWeek = false
    var errorMessage: String?

    var canSubmit: Bool {
        (1...2).contains(selectedDinnerIDs.count) &&
        !isSubmitting &&
        !isCheckingExistingSubmission &&
        !hasExistingSubmissionForWeek
    }

    func loadInitialState(userId: String, weekStartDate: Date) async {
        await fetchDinners()
        await checkExistingSubmission(userId: userId, weekStartDate: weekStartDate)
    }

    func fetchDinners() async {
        isLoadingDinners = true
        errorMessage = nil
        defer { isLoadingDinners = false }

        do {
            let snapshot = try await db
                .collection("dinners")
                .order(by: "name")
                .getDocuments()

            dinners = snapshot.documents.map { document in
                DinnerIdea(
                    id: document.documentID,
                    name: document.data()["name"] as? String ?? "Untitled Dinner",
                    description: document.data()["description"] as? String
                )
            }
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func checkExistingSubmission(userId: String, weekStartDate: Date) async {
        isCheckingExistingSubmission = true
        errorMessage = nil
        defer { isCheckingExistingSubmission = false }

        do {
            let existing = try await db
                .collection("submissions")
                .whereField("userId", isEqualTo: userId)
                .whereField("weekStart", isEqualTo: Timestamp(date: weekStartDate))
                .limit(to: 1)
                .getDocuments()

            hasExistingSubmissionForWeek = !existing.documents.isEmpty
            if hasExistingSubmissionForWeek {
                selectedDinnerIDs = []
            }
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func toggleSelection(for dinnerID: String) {
        if selectedDinnerIDs.contains(dinnerID) {
            selectedDinnerIDs.remove(dinnerID)
            return
        }

        guard selectedDinnerIDs.count < 2 else {
            errorMessage = "You can select up to 2 dinner choices."
            return
        }

        selectedDinnerIDs.insert(dinnerID)
    }

    func submit(userId: String, weekStartDate: Date) async -> Bool {
        guard canSubmit else {
            if hasExistingSubmissionForWeek {
                errorMessage = "You already submitted choices for this week."
            } else if selectedDinnerIDs.isEmpty {
                errorMessage = "Select at least 1 dinner choice."
            } else if selectedDinnerIDs.count > 2 {
                errorMessage = "You can select up to 2 dinner choices."
            }
            return false
        }

        // Check one more time to avoid duplicate writes if state changed remotely.
        await checkExistingSubmission(userId: userId, weekStartDate: weekStartDate)
        guard !hasExistingSubmissionForWeek else {
            errorMessage = "You already submitted choices for this week."
            return false
        }

        isSubmitting = true
        errorMessage = nil
        defer { isSubmitting = false }

        do {
            let choices = Array(selectedDinnerIDs).sorted()
            try await db.collection("submissions").addDocument(data: [
                "userId": userId,
                "weekStart": Timestamp(date: weekStartDate),
                "choices": choices
            ])
            hasExistingSubmissionForWeek = true
            selectedDinnerIDs = []
            return true
        } catch {
            errorMessage = error.localizedDescription
            return false
        }
    }
}
