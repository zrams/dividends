# FamilyDinnerPlanner Cloud Functions

## Scheduled Reminder Function

- Name: `sendWeeklyDinnerChoiceReminders`
- Schedule: `0 17 * * 5`
- Timezone: `America/New_York`
- Trigger: Friday 5:00 PM ET

This function:

1. Finds users with `role == "member"`.
2. Computes next Monday date.
3. Checks `submissions` for members who already submitted.
4. Sends FCM reminders to members missing a submission.
5. Includes deep-link payload to:
   - `familydinnerplanner://member-submission?weekStart=YYYY-MM-DD`

## Local commands

```bash
npm install
npm run lint
npm run serve
```

## Deploy

```bash
npm run deploy
```
