import Foundation
import FirebaseFirestore
import Observation

@MainActor
@Observable
final class AdminDashboardViewModel {
    struct FamilyChoicesByUser: Identifiable, Hashable {
        let id: String
        let userName: String
        let choiceNames: [String]
    }

    private let db = Firestore.firestore()
    private var dinnersListener: ListenerRegistration?
    private var submissionsListener: ListenerRegistration?
    private var userListeners: [String: ListenerRegistration] = [:]
    private var submissionChoicesByUser: [String: [String]] = [:]
    private var userDisplayNames: [String: String] = [:]
    private var selectedFamilyId: String?
    private let lazyPageSize = 25
    private var visibleCount = 25

    var allDinners: [DinnerIdea] = []
    var searchText = ""
    var addName = ""
    var addDescription = ""
    var isLoading = false
    var isSaving = false
    var selectedWeekStartDate = Date.nextMonday(after: .now)
    var isLoadingFamilyChoices = false
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

    var familyChoicesByUser: [FamilyChoicesByUser] {
        let dinnerNamesByID = Dictionary(uniqueKeysWithValues: allDinners.map { ($0.id, $0.name) })

        return submissionChoicesByUser
            .map { userID, rawChoices in
                let uniqueChoices = Self.uniqueValues(from: rawChoices)
                let readableChoices = uniqueChoices.map { dinnerNamesByID[$0] ?? $0 }
                let resolvedName = userDisplayNames[userID] ?? userID
                return FamilyChoicesByUser(
                    id: userID,
                    userName: resolvedName,
                    choiceNames: readableChoices
                )
            }
            .sorted { lhs, rhs in
                lhs.userName.localizedCaseInsensitiveCompare(rhs.userName) == .orderedAscending
            }
    }

    func startListening(familyId: String?) {
        let normalizedFamilyId = normalizedFamilyId(from: familyId)
        let didChangeFamily = normalizedFamilyId != selectedFamilyId
        selectedFamilyId = normalizedFamilyId

        startDinnerIdeasListenerIfNeeded()
        if submissionsListener == nil || didChangeFamily {
            startSubmissionsListener(
                for: selectedWeekStartDate,
                familyId: normalizedFamilyId
            )
        }
    }

    func stopListening() {
        dinnersListener?.remove()
        dinnersListener = nil

        submissionsListener?.remove()
        submissionsListener = nil

        for listener in userListeners.values {
            listener.remove()
        }
        userListeners.removeAll()
        userDisplayNames.removeAll()
        submissionChoicesByUser.removeAll()
        selectedFamilyId = nil
    }

    func updateSelectedWeekStart(_ date: Date) {
        let normalizedDate = Date.mondayForWeek(containing: date)
        guard !Calendar.current.isDate(normalizedDate, inSameDayAs: selectedWeekStartDate) else {
            return
        }

        selectedWeekStartDate = normalizedDate
        startSubmissionsListener(
            for: normalizedDate,
            familyId: selectedFamilyId
        )
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

    private func startDinnerIdeasListenerIfNeeded() {
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

    private func startSubmissionsListener(
        for weekStartDate: Date,
        familyId: String?
    ) {
        submissionsListener?.remove()
        submissionsListener = nil

        isLoadingFamilyChoices = true
        errorMessage = nil

        let weekTimestamp = Timestamp(date: weekStartDate)

        var submissionsQuery: Query = db
            .collection("submissions")
            .whereField("weekStart", isEqualTo: weekTimestamp)

        if let familyId, !familyId.isEmpty {
            submissionsQuery = submissionsQuery.whereField("familyId", isEqualTo: familyId)
        }

        submissionsListener = submissionsQuery
            .addSnapshotListener { [weak self] snapshot, error in
                Task { @MainActor [weak self] in
                    guard let self else { return }

                    if let error {
                        self.errorMessage = error.localizedDescription
                        self.isLoadingFamilyChoices = false
                        return
                    }

                    var groupedChoices: [String: [String]] = [:]
                    for document in snapshot?.documents ?? [] {
                        let data = document.data()
                        guard let userID = data["userId"] as? String else { continue }
                        let choices = data["choices"] as? [String] ?? []
                        groupedChoices[userID, default: []].append(contentsOf: choices)
                    }

                    self.submissionChoicesByUser = groupedChoices
                    self.syncUserListeners(for: Set(groupedChoices.keys))
                    self.isLoadingFamilyChoices = false
                }
            }
    }

    private func normalizedFamilyId(from familyId: String?) -> String? {
        let normalized = familyId?.trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
        return normalized.isEmpty ? nil : normalized
    }

    private func syncUserListeners(for userIDs: Set<String>) {
        let existingUserIDs = Set(userListeners.keys)
        let idsToRemove = existingUserIDs.subtracting(userIDs)
        let idsToAdd = userIDs.subtracting(existingUserIDs)

        for userID in idsToRemove {
            userListeners[userID]?.remove()
            userListeners[userID] = nil
            userDisplayNames[userID] = nil
        }

        for userID in idsToAdd {
            userDisplayNames[userID] = userID

            userListeners[userID] = db.collection("users").document(userID).addSnapshotListener { [weak self] snapshot, error in
                Task { @MainActor [weak self] in
                    guard let self else { return }

                    if let error {
                        self.errorMessage = error.localizedDescription
                        self.userDisplayNames[userID] = userID
                        return
                    }

                    self.userDisplayNames[userID] = self.resolveUserName(
                        from: snapshot?.data(),
                        fallback: userID
                    )
                }
            }
        }
    }

    private func resolveUserName(from data: [String: Any]?, fallback: String) -> String {
        let candidates: [String?] = [
            data?["name"] as? String,
            data?["displayName"] as? String,
            data?["email"] as? String
        ]

        for candidate in candidates {
            let normalized = candidate?.trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
            if !normalized.isEmpty {
                return normalized
            }
        }

        return fallback
    }

    private static func uniqueValues(from values: [String]) -> [String] {
        var seen: Set<String> = []
        return values.filter { value in
            seen.insert(value).inserted
        }
    }
}

private extension Date {
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

    static func mondayForWeek(containing date: Date) -> Date {
        let calendar = Calendar.current
        var components = calendar.dateComponents([.yearForWeekOfYear, .weekOfYear], from: date)
        components.weekday = 2
        return calendar.date(from: components) ?? calendar.startOfDay(for: date)
    }
}
