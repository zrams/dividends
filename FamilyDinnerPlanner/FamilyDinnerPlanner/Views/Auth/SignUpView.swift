import SwiftUI
import Observation

struct SignUpView: View {
    @Bindable var authService: AuthService
    @Environment(\.dismiss) private var dismiss
    let prefilledInviteCode: String? = nil

    @State private var name = ""
    @State private var email = ""
    @State private var password = ""
    @State private var confirmPassword = ""
    @State private var inviteCode = ""
    @State private var localErrorMessage: String?
    @State private var showValidationAlert = false
    @State private var isCreatingAccount = false

    private var showAuthErrorAlert: Binding<Bool> {
        Binding(
            get: { authService.authErrorMessage != nil },
            set: { isPresented in
                if !isPresented {
                    authService.authErrorMessage = nil
                }
            }
        )
    }

    var body: some View {
        Form {
            Section("Create Account") {
                TextField("Name (optional)", text: $name)
                    .textInputAutocapitalization(.words)

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

            Section("Family Invite") {
                TextField("Invite code (optional)", text: $inviteCode)
                    .textInputAutocapitalization(.characters)
                    .autocorrectionDisabled(true)
                Text("If your family admin sent a code, add it here to join the same family.")
                    .font(.footnote)
                    .foregroundStyle(.secondary)
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
        .tint(AppTheme.accent)
        .onAppear {
            if inviteCode.isEmpty, let prefilledInviteCode {
                inviteCode = prefilledInviteCode
            }
        }
        .alert("Sign Up Validation", isPresented: $showValidationAlert) {
            Button("OK", role: .cancel) {
                localErrorMessage = nil
            }
        } message: {
            Text(localErrorMessage ?? "Please check your details and try again.")
        }
        .alert(
            authFailureTitle,
            isPresented: showAuthErrorAlert,
            actions: {
                Button("OK", role: .cancel) {
                    authService.authErrorMessage = nil
                }
            },
            message: {
                Text(authService.authErrorMessage ?? "Please try again.")
            }
        )
    }

    private func createAccount() {
        localErrorMessage = nil

        guard password == confirmPassword else {
            localErrorMessage = "Passwords do not match."
            showValidationAlert = true
            return
        }

        isCreatingAccount = true

        Task {
            defer { isCreatingAccount = false }

            let success = await authService.signUp(
                email: email.trimmingCharacters(in: .whitespacesAndNewlines),
                password: password,
                name: name,
                inviteCode: inviteCode
            )

            if success {
                dismiss()
            }
        }
    }

    private var authFailureTitle: String {
        guard let message = authService.authErrorMessage?.lowercased() else {
            return "Sign Up Failed"
        }
        if message.contains("network") || message.contains("internet") {
            return "Network Error"
        }
        return "Sign Up Failed"
    }
}
