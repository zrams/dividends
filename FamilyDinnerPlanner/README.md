# FamilyDinnerPlanner (iOS 18+, SwiftUI, Firebase)

Native iPhone SwiftUI app scaffold with:

- Firebase initialization in `@main` app entry point
- Email/password auth (login + sign up)
- Role loading from Firestore user profile (`users/{uid}.role`)
- Auth-gated launch routing (`LoginView` vs `MainTabView`)
- Tab-based navigation
- Admin dashboard with searchable dinner management
- Member submission flow with duplicate-week protection
- Admin family choices view with live weekly submission updates
- Admin family management screen with invite flow
- Member submission history for past weeks
- Firestore-backed dinner ideas + weekly submission flow
- Firebase Messaging hooks for push notifications

## Project Structure

```text
FamilyDinnerPlanner/
├── .gitignore
├── README.md
├── firebase/
│   └── functions/
│       ├── index.js
│       ├── package.json
│       └── package-lock.json
└── FamilyDinnerPlanner/
    ├── AppDelegate.swift
    ├── FamilyDinnerPlannerApp.swift
    ├── Config/
    │   ├── FamilyDinnerPlanner.entitlements
    │   └── Info.plist
    ├── Models/
    │   ├── DinnerIdea.swift
    │   ├── FamilyInvite.swift
    │   ├── FamilyMember.swift
    │   ├── MemberSubmissionHistoryItem.swift
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
    ├── ViewModels/
    │   ├── AdminDashboardViewModel.swift
    │   ├── FamilyManagementViewModel.swift
    │   ├── MemberHistoryViewModel.swift
    │   └── MemberSubmissionViewModel.swift
    ├── Theme/
    │   └── AppTheme.swift
    └── Views/
        ├── Auth/
        │   ├── LoginView.swift
        │   └── SignUpView.swift
        ├── Launch/
        │   └── LaunchRouterView.swift
        └── Main/
            ├── AdminDashboardView.swift
            ├── FamilyManagementView.swift
            ├── HomeView.swift
            ├── MainTabView.swift
            ├── MemberHistoryView.swift
            ├── MemberSubmissionView.swift
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

### App-side FCM behavior in this project

- Notification permission requested after login (`NotificationPermissionService`).
- If already authorized, app re-registers for remote notifications at launch/session.
- FCM registration token is captured via `MessagingDelegate`.
- Token is saved to Firestore at `users/{uid}.fcmToken`.
- Push payload supports deep link:
  - `familydinnerplanner://member-submission?weekStart=YYYY-MM-DD`
  - App routes that link to the member submission tab.

## Firestore Data Shape (Starter)

### Collection: `users`

Document ID = Firebase Auth UID

- `role: String` (`admin` or `member`)
- `name: String` (optional display name for admin Family Choices view)
- `fcmToken: String` (saved from iOS app for push notifications)
- `familyId: String` (groups users into a family)

### Collection: `dinners`

Documents can include:

- `name: String`
- `description: String` (optional)

### Collection: `weeklySubmissions`

Saved by `FirestoreService.submitWeeklySubmission(...)`:

- `id: String`
- `userId: String`
- `weekStart: Timestamp`
- `choices: [String]` (array of dinner IDs)

### Collection: `submissions`

Saved by `MemberSubmissionView` for `role == member`:

- `userId: String`
- `weekStart: Timestamp` (next Monday)
- `weekStartISO: String` (`yyyy-MM-dd`)
- `choices: [String]` (selected dinner document IDs)

### Collection: `invites`

Created by admin users in Family Management:

- `code: String` (invite code entered during signup)
- `familyId: String`
- `inviteEmail: String`
- `inviteEmailLowercase: String`
- `inviteLink: String` (custom app link with invite code)
- `status: String` (`active` or `claimed`)
- `createdBy: String`
- `createdAt: Timestamp`
- `claimedBy: String` (set when claimed)
- `claimedAt: Timestamp` (set when claimed)

## Scheduled Reminder Cloud Function (Friday 5 PM EST)

Cloud Function source is in:

- `firebase/functions/index.js`

Function:

- `sendWeeklyDinnerChoiceReminders`
- Schedule: `0 17 * * 5`
- Timezone: `America/New_York` (5 PM Friday in ET/EST/EDT)
- Logic:
  1. Query `users` where `role == member`
  2. Compute next Monday date
  3. Query `submissions` for that week
  4. For members without a submission, send FCM reminder
  5. Personalized message:
     - Title: `Dinner Choices Time!`
     - Body: `Hi [Name], pick your top 1-2 dinners for the week of [Date]!`
  6. Includes deep link payload to member submission screen

### Deploy Cloud Function

From `FamilyDinnerPlanner/firebase/functions`:

1. Install Firebase CLI (if needed): `npm i -g firebase-tools`
2. Authenticate: `firebase login`
3. Initialize project (once): `firebase init functions`
4. Deploy reminder function:
   - `npm run deploy`

> `onSchedule` uses Cloud Scheduler under the hood; ensure billing is enabled for scheduled functions.

## Firebase Console Setup for Push (APNs + FCM)

1. Open **Firebase Console > Project Settings > Cloud Messaging**
2. Under iOS app configuration:
   - Upload APNs Authentication Key (`.p8`)
   - Enter Key ID and Team ID
3. Confirm your iOS app bundle ID matches Firebase app config.
4. In **Apple Developer**:
   - Push Notifications enabled for the App ID
   - APNs key is active
5. In Firestore, each user document should contain:
   - `role`
   - `fcmToken` (automatically written by app after sign-in + permission)

## Auth Flow

- `LaunchRouterView` shows loading state while Firebase Auth resolves session.
- If signed out: user is routed to `LoginView`.
- If signed in: user is routed to `MainTabView`.
- After sign-in, `AuthService` listens to `users/{uid}` and maps `role`.
- `AdminDashboardView` tab is shown only when `role == admin`.
- `MemberSubmissionView` tab is shown only when `role == member`.
- Signup accepts an optional invite code and claims `familyId` from `invites`.

## Admin Dashboard

`AdminDashboardView` includes:

- Live Firestore listener on `dinners` ordered by `name`
- `.searchable()` filter by name/description
- Add dinner form (name required, description optional)
- Per-row Edit/Delete actions
- Lazy list expansion in chunks for larger dinner lists
- "Family Choices" panel with DatePicker for week selection (defaults to upcoming Monday)
- Live Firestore listener on `submissions` filtered by selected `weekStart`
- Grouped display by family member (`users/{uid}` lookup for names)
- Choice ID to dinner-name mapping via the `dinners` listener
- Loading and error states (`ProgressView`, alerts)

Admin write operations are guarded in both UI and view-model methods.

## Family Management (Admin)

`FamilyManagementView` includes:

- Live member list from `users` filtered by `familyId`
- Name + email display for each member
- Invite creation by email (writes to `invites`)
- Generated invite code/link, shareable in-app

## Member History

`MemberHistoryView` includes:

- Past submission history for the signed-in member
- Dinner ID -> dinner name lookup for readable history
- Live updates when submissions change

## Offline and iPhone optimization

- Firestore persistence is explicitly enabled at app launch.
- Dinner-loading flows attempt cache fallback when offline.
- Listener-based screens can render cached data if available.
- App is configured for iPhone-only UI (portrait-first orientation and iPhone device family).

For true enforcement across all clients, configure Firestore Security Rules so only admins can write `dinners`, for example:

```text
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    function isAdmin() {
      return request.auth != null &&
        get(/databases/$(database)/documents/users/$(request.auth.uid)).data.role == "admin";
    }

    match /dinners/{dinnerId} {
      allow read: if request.auth != null;
      allow create, update, delete: if isAdmin();
    }
  }
}
```

