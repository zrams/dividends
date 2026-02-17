const express = require("express");
const Dinner = require("../models/Dinner");
const { authRequired, requireRole } = require("../middleware/auth");

const router = express.Router();

router.get("/", authRequired, async (req, res) => {
  try {
    const dinners = await Dinner.find().sort({ createdAt: -1 });
    return res.json(dinners);
  } catch (error) {
    return res.status(500).json({ message: "Failed to fetch dinner ideas.", error: error.message });
  }
});

router.post("/", authRequired, requireRole("admin"), async (req, res) => {
  try {
    const { name, description } = req.body;
    if (!name || !name.trim()) {
      return res.status(400).json({ message: "Dinner name is required." });
    }

    const dinner = await Dinner.create({
      name: name.trim(),
      description: description ? description.trim() : ""
    });

    return res.status(201).json(dinner);
  } catch (error) {
    return res.status(500).json({ message: "Failed to create dinner idea.", error: error.message });
  }
});

module.exports = router;
