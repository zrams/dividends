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

        let dinnersQuery = db
            .collection("dinners")
            .order(by: "name")

        do {
            let snapshot = try await dinnersQuery.getDocuments()
            dinners = mapDinnerIdeas(from: snapshot.documents)
        } catch {
            do {
                let cachedSnapshot = try await dinnersQuery.getDocuments(source: .cache)
                dinners = mapDinnerIdeas(from: cachedSnapshot.documents)
                if !dinners.isEmpty {
                    errorMessage = "You're offline. Showing cached dinner ideas."
                } else {
                    errorMessage = error.localizedDescription
                }
            } catch {
                errorMessage = error.localizedDescription
            }
        }
    }

    func checkExistingSubmission(userId: String, weekStartDate: Date) async {
        isCheckingExistingSubmission = true
        errorMessage = nil
        defer { isCheckingExistingSubmission = false }

        let existingQuery = db
            .collection("submissions")
            .whereField("userId", isEqualTo: userId)
            .whereField("weekStart", isEqualTo: Timestamp(date: weekStartDate))
            .limit(to: 1)

        do {
            let existing = try await existingQuery.getDocuments()
            hasExistingSubmissionForWeek = !existing.documents.isEmpty
            if hasExistingSubmissionForWeek {
                selectedDinnerIDs = []
            }
        } catch {
            do {
                let cachedExisting = try await existingQuery.getDocuments(source: .cache)
                hasExistingSubmissionForWeek = !cachedExisting.documents.isEmpty
                if hasExistingSubmissionForWeek {
                    selectedDinnerIDs = []
                    errorMessage = "Offline mode: using cached submission status."
                } else {
                    errorMessage = "Offline mode: unable to verify prior submission."
                }
            } catch {
                errorMessage = error.localizedDescription
            }
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

    func submit(
        userId: String,
        weekStartDate: Date,
        familyId: String?
    ) async -> Bool {
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
            var payload: [String: Any] = [
                "userId": userId,
                "weekStart": Timestamp(date: weekStartDate),
                "weekStartISO": Self.weekStartISO(from: weekStartDate),
                "choices": choices
            ]

            if let familyId, !familyId.isEmpty {
                payload["familyId"] = familyId
            }

            try await db.collection("submissions").addDocument(data: payload)
            hasExistingSubmissionForWeek = true
            selectedDinnerIDs = []
            return true
        } catch {
            errorMessage = error.localizedDescription
            return false
        }
    }

    private static func weekStartISO(from date: Date) -> String {
        let components = Calendar.current.dateComponents([.year, .month, .day], from: date)
        let year = components.year ?? 1970
        let month = components.month ?? 1
        let day = components.day ?? 1
        return String(format: "%04d-%02d-%02d", year, month, day)
    }

    private func mapDinnerIdeas(from documents: [QueryDocumentSnapshot]) -> [DinnerIdea] {
        documents.map { document in
            DinnerIdea(
                id: document.documentID,
                name: document.data()["name"] as? String ?? "Untitled Dinner",
                description: document.data()["description"] as? String
            )
        }
    }
}
