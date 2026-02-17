import Foundation
import FirebaseFirestore
import Observation

@MainActor
@Observable
final class AdminDashboardViewModel {
    private let db = Firestore.firestore()
    private var dinnersListener: ListenerRegistration?
    private let lazyPageSize = 25
    private var visibleCount = 25

    var allDinners: [DinnerIdea] = []
    var searchText = ""
    var addName = ""
    var addDescription = ""
    var isLoading = false
    var isSaving = false
    var errorMessage: String?

    var visibleDinners: [DinnerIdea] {
        let query = searchText.trimmingCharacters(in: .whitespacesAndNewlines)

        if query.isEmpty {
            return Array(allDinners.prefix(visibleCount))
        }

        let normalizedQuery = query.lowercased()
        return allDinners.filter { dinner in
            dinner.name.lowercased().contains(normalizedQuery) ||
            (dinner.description?.lowercased().contains(normalizedQuery) ?? false)
        }
    }

    var canLoadMore: Bool {
        searchText.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty &&
        visibleCount < allDinners.count
    }

    func startListening() {
        guard dinnersListener == nil else { return }

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
                        self.isLoading = false
                        return
                    }

                    let documents = snapshot?.documents ?? []
                    self.allDinners = documents.map { document in
                        DinnerIdea(
                            id: document.documentID,
                            name: document.data()["name"] as? String ?? "Untitled Dinner",
                            description: document.data()["description"] as? String
                        )
                    }

                    let minimumVisible = min(self.lazyPageSize, self.allDinners.count)
                    if self.visibleCount < minimumVisible {
                        self.visibleCount = minimumVisible
                    }
                    if self.visibleCount > self.allDinners.count {
                        self.visibleCount = self.allDinners.count
                    }

                    self.isLoading = false
                }
            }
    }

    func stopListening() {
        dinnersListener?.remove()
        dinnersListener = nil
    }

    func loadMoreIfNeeded(currentItem: DinnerIdea?) {
        guard canLoadMore else { return }

        if currentItem == nil || currentItem?.id == visibleDinners.last?.id {
            visibleCount = min(visibleCount + lazyPageSize, allDinners.count)
        }
    }

    func addDinner(isAdmin: Bool) async -> Bool {
        guard isAdmin else {
            errorMessage = "Only admins can add dinner ideas."
            return false
        }

        let name = addName.trimmingCharacters(in: .whitespacesAndNewlines)
        let description = addDescription.trimmingCharacters(in: .whitespacesAndNewlines)

        guard !name.isEmpty else {
            errorMessage = "Dinner name is required."
            return false
        }

        isSaving = true
        errorMessage = nil
        defer { isSaving = false }

        var payload: [String: Any] = [
            "name": name,
            "createdAt": FieldValue.serverTimestamp(),
            "updatedAt": FieldValue.serverTimestamp()
        ]

        if !description.isEmpty {
            payload["description"] = description
        }

        do {
            _ = try await db.collection("dinners").addDocument(data: payload)
            addName = ""
            addDescription = ""
            return true
        } catch {
            errorMessage = error.localizedDescription
            return false
        }
    }

    func updateDinner(
        dinnerID: String,
        name: String,
        description: String,
        isAdmin: Bool
    ) async -> Bool {
        guard isAdmin else {
            errorMessage = "Only admins can edit dinner ideas."
            return false
        }

        let normalizedName = name.trimmingCharacters(in: .whitespacesAndNewlines)
        let normalizedDescription = description.trimmingCharacters(in: .whitespacesAndNewlines)

        guard !normalizedName.isEmpty else {
            errorMessage = "Dinner name is required."
            return false
        }

        isSaving = true
        errorMessage = nil
        defer { isSaving = false }

        var updatePayload: [String: Any] = [
            "name": normalizedName,
            "updatedAt": FieldValue.serverTimestamp()
        ]

        if normalizedDescription.isEmpty {
            updatePayload["description"] = FieldValue.delete()
        } else {
            updatePayload["description"] = normalizedDescription
        }

        do {
            try await db.collection("dinners").document(dinnerID).updateData(updatePayload)
            return true
        } catch {
            errorMessage = error.localizedDescription
            return false
        }
    }

    func deleteDinner(dinnerID: String, isAdmin: Bool) async -> Bool {
        guard isAdmin else {
            errorMessage = "Only admins can delete dinner ideas."
            return false
        }

        isSaving = true
        errorMessage = nil
        defer { isSaving = false }

        do {
            try await db.collection("dinners").document(dinnerID).delete()
            return true
        } catch {
            errorMessage = error.localizedDescription
            return false
        }
    }
}
