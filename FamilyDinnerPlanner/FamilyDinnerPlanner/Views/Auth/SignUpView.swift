import SwiftUI
import Observation

struct SignUpView: View {
    @Bindable var authService: AuthService
    @Environment(\.dismiss) private var dismiss

    @State private var email = ""
    @State private var password = ""
    @State private var confirmPassword = ""
    @State private var localErrorMessage: String?
    @State private var isCreatingAccount = false

    var body: some View {
        Form {
            Section("Create Account") {
                TextField("Email", text: $email)
                    .textContentType(.emailAddress)
                    .keyboardType(.emailAddress)
                    .textInputAutocapitalization(.never)
                    .autocorrectionDisabled(true)

                SecureField("Password", text: $password)
                    .textContentType(.newPassword)

                SecureField("Confirm Password", text: $confirmPassword)
                    .textContentType(.newPassword)
            }

            if let localErrorMessage {
                Section {
                    Text(localErrorMessage)
                        .foregroundStyle(.red)
                }
            } else if let authErrorMessage = authService.authErrorMessage {
                Section {
                    Text(authErrorMessage)
                        .foregroundStyle(.red)
                }
            }

            Section {
                Button {
                    createAccount()
                } label: {
                    if isCreatingAccount {
                        ProgressView()
                            .frame(maxWidth: .infinity)
                    } else {
                        Text("Sign Up")
                            .frame(maxWidth: .infinity)
                    }
                }
                .disabled(
                    isCreatingAccount ||
                    email.isEmpty ||
                    password.isEmpty ||
                    confirmPassword.isEmpty
                )
            }
        }
        .navigationTitle("Sign Up")
        .navigationBarTitleDisplayMode(.inline)
    }

    private func createAccount() {
        localErrorMessage = nil

        guard password == confirmPassword else {
            localErrorMessage = "Passwords do not match."
            return
        }

        isCreatingAccount = true

        Task {
            defer { isCreatingAccount = false }

            let success = await authService.signUp(
                email: email.trimmingCharacters(in: .whitespacesAndNewlines),
                password: password
            )

            if success {
                dismiss()
            }
        }
    }
}
