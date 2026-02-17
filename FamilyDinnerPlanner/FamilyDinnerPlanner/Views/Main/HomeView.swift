import SwiftUI
import Observation

struct HomeView: View {
    @Bindable var firestoreService: FirestoreService

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

    var body: some View {
        Group {
            if firestoreService.isLoadingDinnerIdeas && firestoreService.dinnerIdeas.isEmpty {
                ProgressView("Loading dinner ideas...")
            } else if firestoreService.dinnerIdeas.isEmpty {
                ContentUnavailableView(
                    "No Dinner Ideas Yet",
                    systemImage: "fork.knife",
                    description: Text("Add documents to the 'dinners' Firestore collection.")
                )
            } else {
                List(firestoreService.dinnerIdeas) { dinnerIdea in
                    VStack(alignment: .leading, spacing: 6) {
                        Text(dinnerIdea.name)
                            .font(.headline)
                        if let description = dinnerIdea.description, !description.isEmpty {
                            Text(description)
                                .font(.subheadline)
                                .foregroundStyle(.secondary)
                        }
                    }
                    .padding(.vertical, 4)
                }
                .listStyle(.insetGrouped)
            }
        }
        .navigationTitle("Dinner Ideas")
        .task {
            if firestoreService.dinnerIdeas.isEmpty {
                await firestoreService.fetchDinnerIdeas()
            }
        }
        .refreshable {
            await firestoreService.fetchDinnerIdeas()
        }
        .alert(
            "Unable to Load Ideas",
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
}
