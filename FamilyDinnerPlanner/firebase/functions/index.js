"use strict";

const {onSchedule} = require("firebase-functions/v2/scheduler");
const logger = require("firebase-functions/logger");
const admin = require("firebase-admin");

if (!admin.apps.length) {
  admin.initializeApp();
}

const db = admin.firestore();

exports.sendWeeklyDinnerChoiceReminders = onSchedule(
    {
      schedule: "0 17 * * 5", // Friday 5:00 PM
      timeZone: "America/New_York",
      region: "us-central1",
    },
    async () => {
      const timeZone = "America/New_York";
      const nextMondayISO = getNextMondayISO(timeZone);
      const nextMondayLabel = formatISODate(nextMondayISO, timeZone);
      const deepLink = `familydinnerplanner://member-submission?weekStart=${encodeURIComponent(nextMondayISO)}`;

      logger.info("Running weekly reminder job", {nextMondayISO});

      const membersSnapshot = await db
          .collection("users")
          .where("role", "==", "member")
          .get();

      if (membersSnapshot.empty) {
        logger.info("No member users found.");
        return;
      }

      const submittedUserIDs = await getSubmittedUserIDs(nextMondayISO);

      let sentCount = 0;
      let skippedNoTokenCount = 0;
      let skippedAlreadySubmittedCount = 0;
      let failedCount = 0;

      for (const memberDocument of membersSnapshot.docs) {
        const userID = memberDocument.id;
        const memberData = memberDocument.data() || {};

        if (submittedUserIDs.has(userID)) {
          skippedAlreadySubmittedCount += 1;
          continue;
        }

        const fcmToken = memberData.fcmToken;
        if (!fcmToken || typeof fcmToken !== "string") {
          skippedNoTokenCount += 1;
          continue;
        }

        const memberName = resolveMemberName(memberData);
        const body =
          `Hi ${memberName}, pick your top 1-2 dinners for the week of ${nextMondayLabel}!`;

        const message = {
          token: fcmToken,
          notification: {
            title: "Dinner Choices Time!",
            body,
          },
          data: {
            deepLink,
            screen: "member-submission",
            weekStartISO: nextMondayISO,
            notificationType: "weekly-dinner-reminder",
          },
          apns: {
            payload: {
              aps: {
                sound: "default",
                category: "WEEKLY_DINNER_REMINDER",
              },
            },
          },
        };

        try {
          await admin.messaging().send(message);
          sentCount += 1;
        } catch (error) {
          failedCount += 1;
          logger.error("Failed to send reminder", {userID, error});

          if (
            error &&
            (error.code === "messaging/registration-token-not-registered" ||
              error.code === "messaging/invalid-registration-token")
          ) {
            await db.collection("users").doc(userID).set(
                {fcmToken: admin.firestore.FieldValue.delete()},
                {merge: true},
            );
          }
        }
      }

      logger.info("Weekly reminder job finished", {
        nextMondayISO,
        sentCount,
        skippedNoTokenCount,
        skippedAlreadySubmittedCount,
        failedCount,
      });
    },
);

async function getSubmittedUserIDs(nextMondayISO) {
  const submittedUserIDs = new Set();

  // Preferred query path for current app documents.
  const isoSnapshot = await db
      .collection("submissions")
      .where("weekStartISO", "==", nextMondayISO)
      .get();

  for (const document of isoSnapshot.docs) {
    const userID = document.data().userId;
    if (typeof userID === "string" && userID.length > 0) {
      submittedUserIDs.add(userID);
    }
  }

  // Backward-compatible fallback for documents created before weekStartISO.
  const {startOfDayUTC, nextDayUTC} = isoDateUTCWindow(nextMondayISO);
  const legacySnapshot = await db
      .collection("submissions")
      .where("weekStart", ">=", admin.firestore.Timestamp.fromDate(startOfDayUTC))
      .where("weekStart", "<", admin.firestore.Timestamp.fromDate(nextDayUTC))
      .get();

  for (const document of legacySnapshot.docs) {
    const userID = document.data().userId;
    if (typeof userID === "string" && userID.length > 0) {
      submittedUserIDs.add(userID);
    }
  }

  return submittedUserIDs;
}

function resolveMemberName(memberData) {
  const preferred =
    stringOrNull(memberData.name) ||
    stringOrNull(memberData.displayName) ||
    stringOrNull(memberData.firstName);

  if (preferred) {
    return preferred;
  }

  const email = stringOrNull(memberData.email);
  if (email && email.includes("@")) {
    return email.split("@")[0];
  }

  return "there";
}

function stringOrNull(value) {
  if (typeof value !== "string") return null;
  const normalized = value.trim();
  return normalized.length > 0 ? normalized : null;
}

function getNextMondayISO(timeZone) {
  const now = new Date();
  const {year, month, day, weekdayIndex} = zonedDateParts(now, timeZone);
  const currentDateUTC = new Date(Date.UTC(year, month - 1, day));
  const daysUntilNextMonday = weekdayIndex === 1 ? 7 : (8 - weekdayIndex) % 7;
  const nextMondayUTC = new Date(
      currentDateUTC.getTime() + daysUntilNextMonday * 24 * 60 * 60 * 1000,
  );

  const nextYear = nextMondayUTC.getUTCFullYear();
  const nextMonth = String(nextMondayUTC.getUTCMonth() + 1).padStart(2, "0");
  const nextDay = String(nextMondayUTC.getUTCDate()).padStart(2, "0");
  return `${nextYear}-${nextMonth}-${nextDay}`;
}

function zonedDateParts(date, timeZone) {
  const formatter = new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    weekday: "short",
  });

  const parts = formatter.formatToParts(date);
  const year = Number(parts.find((part) => part.type === "year")?.value);
  const month = Number(parts.find((part) => part.type === "month")?.value);
  const day = Number(parts.find((part) => part.type === "day")?.value);
  const weekdayToken = parts.find((part) => part.type === "weekday")?.value || "Mon";

  return {
    year,
    month,
    day,
    weekdayIndex: weekdayFromToken(weekdayToken),
  };
}

function weekdayFromToken(token) {
  const map = {
    Mon: 1,
    Tue: 2,
    Wed: 3,
    Thu: 4,
    Fri: 5,
    Sat: 6,
    Sun: 7,
  };
  return map[token] || 1;
}

function isoDateUTCWindow(isoDate) {
  const startOfDayUTC = new Date(`${isoDate}T00:00:00.000Z`);
  const nextDayUTC = new Date(startOfDayUTC.getTime() + 24 * 60 * 60 * 1000);
  return {startOfDayUTC, nextDayUTC};
}

function formatISODate(isoDate, timeZone) {
  const date = new Date(`${isoDate}T00:00:00.000Z`);
  return new Intl.DateTimeFormat("en-US", {
    timeZone,
    month: "short",
    day: "numeric",
    year: "numeric",
  }).format(date);
}
