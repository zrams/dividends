import Foundation
import FirebaseFirestore
import Observation

@MainActor
@Observable
final class FirestoreService {
    private enum Collections {
        static let dinners = "dinners"
        static let weeklySubmissions = "weeklySubmissions"
    }

    private let db = Firestore.firestore()

    var dinnerIdeas: [DinnerIdea] = []
    var isLoadingDinnerIdeas = false
    var firestoreErrorMessage: String?

    func fetchDinnerIdeas() async {
        firestoreErrorMessage = nil
        isLoadingDinnerIdeas = true
        defer { isLoadingDinnerIdeas = false }

        do {
            let snapshot = try await db
                .collection(Collections.dinners)
                .order(by: "name")
                .getDocuments()

            dinnerIdeas = snapshot.documents.map { document in
                DinnerIdea(
                    id: document.documentID,
                    name: document.data()["name"] as? String ?? "Untitled Dinner",
                    description: document.data()["description"] as? String
                )
            }
        } catch {
            firestoreErrorMessage = error.localizedDescription
        }
    }

    func submitWeeklySubmission(
        userId: String,
        choices: [String],
        weekStart: Date = .startOfWeek
    ) async -> Bool {
        firestoreErrorMessage = nil

        let documentId = "\(userId)_\(weekStart.yyyyMMdd)"
        let submission = WeeklySubmission(
            id: documentId,
            userId: userId,
            weekStart: weekStart,
            choices: choices
        )

        do {
            try await db
                .collection(Collections.weeklySubmissions)
                .document(documentId)
                .setData(submission.asFirestoreData)
            return true
        } catch {
            firestoreErrorMessage = error.localizedDescription
            return false
        }
    }
}

private extension Date {
    static var startOfWeek: Date {
        let calendar = Calendar(identifier: .gregorian)
        let now = Date()
        let components = calendar.dateComponents([.yearForWeekOfYear, .weekOfYear], from: now)
        return calendar.date(from: components) ?? now
    }

    var yyyyMMdd: String {
        let formatter = DateFormatter()
        formatter.locale = Locale(identifier: "en_US_POSIX")
        formatter.timeZone = TimeZone(secondsFromGMT: 0)
        formatter.dateFormat = "yyyy-MM-dd"
        return formatter.string(from: self)
    }
}
