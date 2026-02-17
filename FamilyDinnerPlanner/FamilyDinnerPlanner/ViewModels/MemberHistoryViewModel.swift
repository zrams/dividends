import Foundation
import FirebaseFirestore
import Observation

@MainActor
@Observable
final class MemberHistoryViewModel {
    private let db = Firestore.firestore()
    private var submissionsListener: ListenerRegistration?
    private var dinnersListener: ListenerRegistration?
    private var submissionDocuments: [QueryDocumentSnapshot] = []
    private var dinnerNamesByID: [String: String] = [:]

    var historyItems: [MemberSubmissionHistoryItem] = []
    var isLoading = false
    var errorMessage: String?

    func startListening(userId: String) {
        guard !userId.isEmpty else { return }
        if submissionsListener != nil || dinnersListener != nil {
            return
        }

        isLoading = true
        errorMessage = nil

        dinnersListener = db
            .collection("dinners")
            .order(by: "name")
            .addSnapshotListener { [weak self] snapshot, error in
                Task { @MainActor [weak self] in
                    guard let self else { return }

                    if let error {
                        self.errorMessage = error.localizedDescription
                        return
                    }

                    let names = (snapshot?.documents ?? []).reduce(into: [String: String]()) { partialResult, document in
                        partialResult[document.documentID] = document.data()["name"] as? String ?? document.documentID
                    }
                    self.dinnerNamesByID = names
                    self.rebuildHistoryItems()
                }
            }

        submissionsListener = db
            .collection("submissions")
            .whereField("userId", isEqualTo: userId)
            .addSnapshotListener { [weak self] snapshot, error in
                Task { @MainActor [weak self] in
                    guard let self else { return }

                    if let error {
                        self.errorMessage = error.localizedDescription
                        self.isLoading = false
                        return
                    }

                    self.submissionDocuments = snapshot?.documents ?? []
                    self.rebuildHistoryItems()
                    self.isLoading = false
                }
            }
    }

    func stopListening() {
        submissionsListener?.remove()
        submissionsListener = nil
        dinnersListener?.remove()
        dinnersListener = nil
    }

    private func rebuildHistoryItems() {
        historyItems = submissionDocuments
            .compactMap { document -> MemberSubmissionHistoryItem? in
                let data = document.data()
                guard let weekStart = parseWeekStart(from: data) else { return nil }
                let choiceIDs = data["choices"] as? [String] ?? []
                let choiceNames = choiceIDs.map { dinnerNamesByID[$0] ?? $0 }

                return MemberSubmissionHistoryItem(
                    id: document.documentID,
                    weekStart: weekStart,
                    choiceNames: choiceNames
                )
            }
            .sorted { lhs, rhs in
                lhs.weekStart > rhs.weekStart
            }
    }

    private func parseWeekStart(from data: [String: Any]) -> Date? {
        if let timestamp = data["weekStart"] as? Timestamp {
            return timestamp.dateValue()
        }

        if let isoString = data["weekStartISO"] as? String {
            let components = isoString.split(separator: "-")
            if components.count == 3,
               let year = Int(components[0]),
               let month = Int(components[1]),
               let day = Int(components[2]) {
                let calendar = Calendar.current
                return calendar.date(from: DateComponents(year: year, month: month, day: day))
            }
        }

        return nil
    }
}
