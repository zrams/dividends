import SwiftUI
import Observation

struct AdminDashboardView: View {
    private enum AdminPanel: String, CaseIterable, Identifiable {
        case dinners = "Dinner Ideas"
        case familyChoices = "Family Choices"

        var id: String { rawValue }
    }

    @Bindable var authService: AuthService
    @State private var viewModel = AdminDashboardViewModel()
    @State private var selectedPanel: AdminPanel = .dinners

    @State private var editingDinner: DinnerIdea?
    @State private var dinnerPendingDelete: DinnerIdea?
    @State private var showDeleteConfirmation = false

    private var isAdmin: Bool {
        authService.isAdmin
    }

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

    private var weekSelectionBinding: Binding<Date> {
        Binding(
            get: { viewModel.selectedWeekStartDate },
            set: { newDate in
                viewModel.updateSelectedWeekStart(newDate)
            }
        )
    }

    var body: some View {
        Group {
            if selectedPanel == .dinners && viewModel.isLoading && viewModel.allDinners.isEmpty {
                ProgressView("Loading dinners...")
            } else if selectedPanel == .familyChoices &&
                        viewModel.isLoadingFamilyChoices &&
                        viewModel.familyChoicesByUser.isEmpty {
                ProgressView("Loading family choices...")
            } else {
                if selectedPanel == .dinners {
                    dashboardList
                        .searchable(text: $viewModel.searchText, prompt: "Search dinners")
                } else {
                    dashboardList
                }
            }
        }
        .navigationTitle("Admin Dashboard")
        .toolbar {
            ToolbarItem(placement: .topBarTrailing) {
                Image(systemName: "flame.fill")
                    .foregroundStyle(AppTheme.accent)
            }
        }
        .task {
            viewModel.startListening(familyId: authService.currentFamilyId)
        }
        .onChange(of: authService.currentFamilyId) { _, newFamilyId in
            viewModel.startListening(familyId: newFamilyId)
        }
        .onDisappear {
            viewModel.stopListening()
        }
        .alert(
            "Admin Action Failed",
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
        .confirmationDialog(
            "Delete this dinner idea?",
            isPresented: $showDeleteConfirmation,
            titleVisibility: .visible
        ) {
            Button("Delete", role: .destructive) {
                deletePendingDinner()
            }
            Button("Cancel", role: .cancel) {
                dinnerPendingDelete = nil
            }
        } message: {
            Text(dinnerPendingDelete?.name ?? "")
        }
        .sheet(item: $editingDinner) { dinner in
            EditDinnerView(
                dinner: dinner,
                isSaving: viewModel.isSaving,
                isAdmin: isAdmin,
                onCancel: {
                    editingDinner = nil
                },
                onSave: { name, description in
                    saveDinnerEdits(dinnerID: dinner.id, name: name, description: description)
                }
            )
        }
    }

    private var dashboardList: some View {
        List {
            Section {
                Picker("Admin Panel", selection: $selectedPanel) {
                    ForEach(AdminPanel.allCases) { panel in
                        Text(panel.rawValue).tag(panel)
                    }
                }
                .pickerStyle(.segmented)
            }

            switch selectedPanel {
            case .dinners:
                dinnerManagementSections
            case .familyChoices:
                familyChoicesSections
            }
        }
        .listStyle(.insetGrouped)
        .tint(AppTheme.accent)
    }

    @ViewBuilder
    private var dinnerManagementSections: some View {
        Section {
            TextField("Dinner name (required)", text: $viewModel.addName)
                .textInputAutocapitalization(.words)

            TextField("Description (optional)", text: $viewModel.addDescription, axis: .vertical)
                .lineLimit(2...5)

            Button {
                Task {
                    _ = await viewModel.addDinner(isAdmin: isAdmin)
                }
            } label: {
                if viewModel.isSaving {
                    ProgressView()
                        .frame(maxWidth: .infinity)
                } else {
                    Text("Add Dinner")
                        .frame(maxWidth: .infinity)
                }
            }
            .disabled(viewModel.isSaving || !isAdmin)
        } header: {
            Label("Add New Dinner", systemImage: "fork.knife.circle")
        }

        Section {
            if viewModel.visibleDinners.isEmpty {
                ContentUnavailableView(
                    "No Dinners Found",
                    systemImage: "fork.knife",
                    description: Text("Add a dinner above or adjust your search query.")
                )
            } else {
                ForEach(viewModel.visibleDinners) { dinner in
                    AdminDinnerRow(
                        dinner: dinner,
                        isAdmin: isAdmin,
                        onEdit: { editingDinner = dinner },
                        onDelete: {
                            dinnerPendingDelete = dinner
                            showDeleteConfirmation = true
                        }
                    )
                    .onAppear {
                        viewModel.loadMoreIfNeeded(currentItem: dinner)
                    }
                }
            }

            if viewModel.canLoadMore {
                HStack {
                    Spacer()
                    ProgressView("Loading more...")
                    Spacer()
                }
                .onAppear {
                    viewModel.loadMoreIfNeeded(currentItem: nil)
                }
            }
        } header: {
            Label("Dinner Ideas (\(viewModel.allDinners.count))", systemImage: "list.bullet.rectangle")
        }
    }

    @ViewBuilder
    private var familyChoicesSections: some View {
        Section {
            DatePicker(
                "Week Starting (Monday)",
                selection: weekSelectionBinding,
                displayedComponents: .date
            )

            Text("Live submissions for week of \(formattedDate(viewModel.selectedWeekStartDate)).")
                .font(.subheadline)
                .foregroundStyle(.secondary)
        } header: {
            Label("Family Choices", systemImage: "person.3")
        }

        Section {
            if viewModel.isLoadingFamilyChoices && viewModel.familyChoicesByUser.isEmpty {
                ProgressView("Loading submissions...")
                    .frame(maxWidth: .infinity, alignment: .center)
            } else if viewModel.familyChoicesByUser.isEmpty {
                ContentUnavailableView(
                    "No Submissions Yet",
                    systemImage: "tray",
                    description: Text("No family member submissions found for this week.")
                )
            } else {
                ForEach(viewModel.familyChoicesByUser) { choiceGroup in
                    FamilyChoicesRow(choiceGroup: choiceGroup)
                }
            }
        } header: {
            Label("Member Submissions", systemImage: "tray.full")
        }
    }

    private func formattedDate(_ date: Date) -> String {
        let formatter = DateFormatter()
        formatter.dateStyle = .medium
        formatter.timeStyle = .none
        return formatter.string(from: date)
    }

    private func saveDinnerEdits(dinnerID: String, name: String, description: String) {
        Task {
            let success = await viewModel.updateDinner(
                dinnerID: dinnerID,
                name: name,
                description: description,
                isAdmin: isAdmin
            )
            if success {
                editingDinner = nil
            }
        }
    }

    private func deletePendingDinner() {
        guard let dinnerPendingDelete else { return }

        Task {
            let success = await viewModel.deleteDinner(
                dinnerID: dinnerPendingDelete.id,
                isAdmin: isAdmin
            )

            if success {
                self.dinnerPendingDelete = nil
            }
        }
    }
}

private struct FamilyChoicesRow: View {
    let choiceGroup: AdminDashboardViewModel.FamilyChoicesByUser

    var body: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text(choiceGroup.userName)
                .font(.headline)

            if choiceGroup.choiceNames.isEmpty {
                Text("No dinner choices submitted.")
                    .font(.subheadline)
                    .foregroundStyle(.secondary)
            } else {
                ForEach(choiceGroup.choiceNames, id: \.self) { choiceName in
                    Text("• \(choiceName)")
                        .font(.subheadline)
                }
            }
        }
        .padding(.vertical, 4)
    }
}

private struct AdminDinnerRow: View {
    let dinner: DinnerIdea
    let isAdmin: Bool
    let onEdit: () -> Void
    let onDelete: () -> Void

    var body: some View {
        HStack(alignment: .top, spacing: 12) {
            VStack(alignment: .leading, spacing: 5) {
                Text(dinner.name)
                    .font(.headline)
                if let description = dinner.description, !description.isEmpty {
                    Text(description)
                        .font(.subheadline)
                        .foregroundStyle(.secondary)
                }
            }

            Spacer(minLength: 8)

            if isAdmin {
                HStack(spacing: 12) {
                    Button("Edit", action: onEdit)
                        .buttonStyle(.borderless)

                    Button("Delete", role: .destructive, action: onDelete)
                        .buttonStyle(.borderless)
                }
                .font(.caption.weight(.semibold))
            }
        }
        .padding(.vertical, 4)
    }
}

private struct EditDinnerView: View {
    let dinner: DinnerIdea
    let isSaving: Bool
    let isAdmin: Bool
    let onCancel: () -> Void
    let onSave: (_ name: String, _ description: String) -> Void

    @State private var name: String
    @State private var description: String

    init(
        dinner: DinnerIdea,
        isSaving: Bool,
        isAdmin: Bool,
        onCancel: @escaping () -> Void,
        onSave: @escaping (_ name: String, _ description: String) -> Void
    ) {
        self.dinner = dinner
        self.isSaving = isSaving
        self.isAdmin = isAdmin
        self.onCancel = onCancel
        self.onSave = onSave
        _name = State(initialValue: dinner.name)
        _description = State(initialValue: dinner.description ?? "")
    }

    var body: some View {
        NavigationStack {
            Form {
                Section("Edit Dinner") {
                    TextField("Dinner name", text: $name)
                        .textInputAutocapitalization(.words)

                    TextField("Description", text: $description, axis: .vertical)
                        .lineLimit(2...5)
                }
            }
            .navigationTitle("Edit")
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .cancellationAction) {
                    Button("Cancel") {
                        onCancel()
                    }
                }

                ToolbarItem(placement: .confirmationAction) {
                    Button {
                        onSave(name, description)
                    } label: {
                        if isSaving {
                            ProgressView()
                        } else {
                            Text("Save")
                        }
                    }
                    .disabled(
                        isSaving ||
                        !isAdmin ||
                        name.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty
                    )
                }
            }
        }
    }
}
