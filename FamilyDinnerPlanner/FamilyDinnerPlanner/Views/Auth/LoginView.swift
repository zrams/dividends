import SwiftUI
import Observation

struct LoginView: View {
    @Bindable var authService: AuthService

    @State private var email = ""
    @State private var password = ""
    @State private var isAuthenticating = false

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

            if let authErrorMessage = authService.authErrorMessage {
                Text(authErrorMessage)
                    .font(.footnote)
                    .foregroundStyle(.red)
                    .multilineTextAlignment(.center)
                    .frame(maxWidth: .infinity, alignment: .leading)
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
            .tint(.orange)
            .disabled(isAuthenticating || email.isEmpty || password.isEmpty)

            NavigationLink {
                SignUpView(authService: authService)
            } label: {
                Text("Create an Account")
            }
            .font(.callout)
            .padding(.top, 6)

            Spacer()
        }
        .padding(24)
        .navigationBarTitleDisplayMode(.inline)
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
}
