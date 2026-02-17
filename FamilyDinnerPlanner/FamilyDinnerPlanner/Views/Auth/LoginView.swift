import SwiftUI
import Observation

struct LoginView: View {
    @Bindable var authService: AuthService

    @State private var email = ""
    @State private var password = ""
    @State private var inviteCodeFromLink = ""
    @State private var isAuthenticating = false

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
        VStack(spacing: 20) {
            Spacer()

            Image(systemName: "fork.knife.circle.fill")
                .font(.system(size: 64))
                .foregroundStyle(.orange)

            Text("FamilyDinnerPlanner")
                .font(.largeTitle.bold())

            VStack(spacing: 12) {
                TextField("Email", text: $email)
                    .textContentType(.emailAddress)
                    .keyboardType(.emailAddress)
                    .textInputAutocapitalization(.never)
                    .autocorrectionDisabled(true)
                    .textFieldStyle(.roundedBorder)

                SecureField("Password", text: $password)
                    .textContentType(.password)
                    .textFieldStyle(.roundedBorder)
            }

            Button {
                signIn()
            } label: {
                if isAuthenticating {
                    ProgressView()
                        .frame(maxWidth: .infinity)
                } else {
                    Text("Log In")
                        .frame(maxWidth: .infinity)
                }
            }
            .buttonStyle(.borderedProminent)
            .tint(AppTheme.accent)
            .disabled(isAuthenticating || email.isEmpty || password.isEmpty)

            NavigationLink {
                SignUpView(
                    authService: authService,
                    prefilledInviteCode: inviteCodeFromLink.isEmpty ? nil : inviteCodeFromLink
                )
            } label: {
                Text("Create an Account")
            }
            .font(.callout)
            .padding(.top, 6)

            Spacer()
        }
        .padding(24)
        .background(
            LinearGradient(
                colors: [AppTheme.backgroundTop, AppTheme.backgroundBottom],
                startPoint: .top,
                endPoint: .bottom
            )
            .ignoresSafeArea()
        )
        .navigationBarTitleDisplayMode(.inline)
        .onOpenURL { url in
            hydrateInviteCode(from: url)
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

    private func signIn() {
        isAuthenticating = true

        Task {
            defer { isAuthenticating = false }
            _ = await authService.signIn(
                email: email.trimmingCharacters(in: .whitespacesAndNewlines),
                password: password
            )
        }
    }

    private func hydrateInviteCode(from url: URL) {
        guard url.scheme?.lowercased() == "familydinnerplanner" else { return }
        guard (url.host ?? "").lowercased() == "signup" || url.path.lowercased().contains("signup") else {
            return
        }

        guard let components = URLComponents(url: url, resolvingAgainstBaseURL: false),
              let inviteCode = components.queryItems?.first(where: { $0.name == "inviteCode" })?.value,
              !inviteCode.isEmpty else {
            return
        }

        inviteCodeFromLink = inviteCode
    }

    private var authFailureTitle: String {
        guard let message = authService.authErrorMessage?.lowercased() else {
            return "Authentication Failed"
        }
        if message.contains("network") || message.contains("internet") {
            return "Network Error"
        }
        return "Authentication Failed"
    }
}
