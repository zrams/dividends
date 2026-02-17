# FamilyDinnerPlanner (iOS 18+, SwiftUI, Firebase)

Native iPhone SwiftUI app scaffold with:

- Firebase initialization in `@main` app entry point
- Email/password auth (login + sign up)
- Auth-gated launch routing (`LoginView` vs `MainTabView`)
- Tab-based navigation
- Firestore-backed dinner ideas + weekly submission flow
- Firebase Messaging hooks for push notifications

## Project Structure

```text
FamilyDinnerPlanner/
├── .gitignore
├── README.md
└── FamilyDinnerPlanner/
    ├── AppDelegate.swift
    ├── FamilyDinnerPlannerApp.swift
    ├── Config/
    │   ├── FamilyDinnerPlanner.entitlements
    │   └── Info.plist
    ├── Models/
    │   ├── DinnerIdea.swift
    │   ├── UserRole.swift
    │   └── WeeklySubmission.swift
    ├── Resources/
    │   ├── Assets.xcassets/
    │   │   ├── AppIcon.appiconset/Contents.json
    │   │   ├── Contents.json
    │   │   └── LaunchBackground.colorset/Contents.json
    │   └── Preview Content/
    │       └── Preview Assets.xcassets/Contents.json
    ├── Services/
    │   ├── AuthService.swift
    │   ├── FirestoreService.swift
    │   └── NotificationPermissionService.swift
    └── Views/
        ├── Auth/
        │   ├── LoginView.swift
        │   └── SignUpView.swift
        ├── Launch/
        │   └── LaunchRouterView.swift
        └── Main/
            ├── HomeView.swift
            ├── MainTabView.swift
            ├── ProfileView.swift
            └── WeeklySubmissionView.swift
```

## Key App Entry Point

`FamilyDinnerPlannerApp.swift` configures Firebase in `init()` and imports:

- `FirebaseCore`
- `FirebaseAuth`
- `FirebaseFirestore`
- `FirebaseMessaging`

## Xcode Setup (Create/Open Project + Add Files)

Because this environment cannot run Xcode, create/open your iOS app target in Xcode and then add the files from `FamilyDinnerPlanner/FamilyDinnerPlanner/` into that target.

Recommended new project settings:

- Platform: iOS
- Interface: SwiftUI
- Language: Swift
- Minimum Deployment: iOS 18.0
- Devices: iPhone

## Add Firebase via Swift Package Manager

1. In Xcode, go to **File > Add Package Dependencies...**
2. Package URL:
   - `https://github.com/firebase/firebase-ios-sdk.git`
3. Dependency rule:
   - **Up to Next Major Version**
4. Add these products to your app target:
   - `FirebaseCore`
   - `FirebaseAuth`
   - `FirebaseFirestore`
   - `FirebaseMessaging`

## Add `GoogleService-Info.plist`

1. In Firebase Console, create/select your iOS app (`Bundle ID` must match Xcode target).
2. Download `GoogleService-Info.plist`.
3. Drag it into Xcode project navigator.
4. Ensure:
   - **Copy items if needed** is checked
   - Your app target is selected in **Target Membership**

## Push Notifications / FCM Capabilities

In Xcode target **Signing & Capabilities**:

1. Add **Push Notifications**
2. Add **Background Modes**
   - Check **Remote notifications**
3. Confirm entitlements include `aps-environment`
4. In Apple Developer portal for your App ID:
   - Enable Push Notifications
5. Upload APNs auth key (`.p8`) in Firebase Console under Cloud Messaging

## Firestore Data Shape (Starter)

### Collection: `dinnerIdeas`

Documents can include:

- `name: String`
- `description: String` (optional)

### Collection: `weeklySubmissions`

Saved by `FirestoreService.submitWeeklySubmission(...)`:

- `id: String`
- `userId: String`
- `weekStart: Timestamp`
- `choices: [String]` (array of dinner IDs)

## Auth Flow

- `LaunchRouterView` shows loading state while Firebase Auth resolves session.
- If signed out: user is routed to `LoginView`.
- If signed in: user is routed to `MainTabView`.

