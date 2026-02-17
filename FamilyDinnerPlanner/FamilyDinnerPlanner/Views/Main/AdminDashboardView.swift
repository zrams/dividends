import SwiftUI
import Observation

struct AdminDashboardView: View {
    @Bindable var authService: AuthService
    @State private var viewModel = AdminDashboardViewModel()

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

    var body: some View {
        Group {
            if viewModel.isLoading && viewModel.allDinners.isEmpty {
                ProgressView("Loading dinners...")
            } else {
                List {
                    Section("Add New Dinner") {
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
                    }

                    Section("Dinner Ideas (\(viewModel.allDinners.count))") {
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
                    }
                }
                .listStyle(.insetGrouped)
                .searchable(text: $viewModel.searchText, prompt: "Search dinners")
            }
        }
        .navigationTitle("Admin Dashboard")
        .task {
            viewModel.startListening()
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
