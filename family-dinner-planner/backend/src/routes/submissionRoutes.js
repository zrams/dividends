const express = require("express");
const mongoose = require("mongoose");
const Dinner = require("../models/Dinner");
const WeeklySubmission = require("../models/WeeklySubmission");
const { authRequired, requireRole } = require("../middleware/auth");

const router = express.Router();

function normalizeWeekStartDate(inputDate) {
  const date = new Date(inputDate);
  if (Number.isNaN(date.getTime())) {
    return null;
  }

  // Normalize to Monday 00:00 UTC for uniqueness per week.
  date.setUTCHours(0, 0, 0, 0);
  const day = date.getUTCDay();
  const diffFromMonday = (day + 6) % 7;
  date.setUTCDate(date.getUTCDate() - diffFromMonday);
  return date;
}

router.post("/", authRequired, requireRole("member"), async (req, res) => {
  try {
    const { weekStartDate, choices } = req.body;

    if (!weekStartDate) {
      return res.status(400).json({ message: "weekStartDate is required." });
    }
    if (!Array.isArray(choices) || choices.length < 1 || choices.length > 2) {
      return res.status(400).json({ message: "choices must contain 1-2 dinner IDs." });
    }

    const normalizedDate = normalizeWeekStartDate(weekStartDate);
    if (!normalizedDate) {
      return res.status(400).json({ message: "weekStartDate must be a valid date string." });
    }

    const uniqueChoices = [...new Set(choices)];
    if (uniqueChoices.length !== choices.length) {
      return res.status(400).json({ message: "Duplicate dinner choices are not allowed." });
    }

    const areValidObjectIds = uniqueChoices.every((id) => mongoose.Types.ObjectId.isValid(id));
    if (!areValidObjectIds) {
      return res.status(400).json({ message: "One or more dinner IDs are invalid." });
    }

    const matchingDinnersCount = await Dinner.countDocuments({
      _id: { $in: uniqueChoices }
    });
    if (matchingDinnersCount !== uniqueChoices.length) {
      return res.status(400).json({ message: "One or more selected dinners do not exist." });
    }

    const submission = await WeeklySubmission.findOneAndUpdate(
      {
        userId: req.user._id,
        weekStartDate: normalizedDate
      },
      {
        userId: req.user._id,
        weekStartDate: normalizedDate,
        choices: uniqueChoices
      },
      {
        new: true,
        upsert: true,
        runValidators: true,
        setDefaultsOnInsert: true
      }
    ).populate("choices", "name description");

    return res.status(201).json(submission);
  } catch (error) {
    return res.status(500).json({ message: "Failed to save weekly submission.", error: error.message });
  }
});

router.get("/mine", authRequired, requireRole("member"), async (req, res) => {
  try {
    const submissions = await WeeklySubmission.find({ userId: req.user._id })
      .sort({ weekStartDate: -1 })
      .populate("choices", "name description");
    return res.json(submissions);
  } catch (error) {
    return res.status(500).json({ message: "Failed to fetch your submissions.", error: error.message });
  }
});

router.get("/", authRequired, requireRole("admin"), async (req, res) => {
  try {
    const { weekStartDate } = req.query;
    const query = {};

    if (weekStartDate) {
      const normalizedDate = normalizeWeekStartDate(weekStartDate);
      if (!normalizedDate) {
        return res.status(400).json({ message: "weekStartDate query must be a valid date string." });
      }
      query.weekStartDate = normalizedDate;
    }

    const submissions = await WeeklySubmission.find(query)
      .sort({ weekStartDate: -1, createdAt: -1 })
      .populate("userId", "name email role")
      .populate("choices", "name description");

    return res.json(submissions);
  } catch (error) {
    return res.status(500).json({ message: "Failed to fetch submissions.", error: error.message });
  }
});

module.exports = router;
