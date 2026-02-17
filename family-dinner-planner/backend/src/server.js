const express = require("express");
const cors = require("cors");
const dotenv = require("dotenv");

const connectDb = require("./config/db");
const authRoutes = require("./routes/authRoutes");
const dinnerRoutes = require("./routes/dinnerRoutes");
const submissionRoutes = require("./routes/submissionRoutes");

dotenv.config();

const app = express();
const port = process.env.PORT || 5000;

app.use(
  cors({
    origin: process.env.CLIENT_URL || "http://localhost:5173"
  })
);
app.use(express.json());

app.get("/api/health", (req, res) => {
  res.json({ status: "ok" });
});

app.use("/api/auth", authRoutes);
app.use("/api/dinners", dinnerRoutes);
app.use("/api/submissions", submissionRoutes);

app.use((req, res) => {
  res.status(404).json({ message: "Route not found." });
});

connectDb()
  .then(() => {
    app.listen(port, () => {
      console.log(`Backend API running on http://localhost:${port}`);
    });
  })
  .catch((error) => {
    console.error("Failed to start server:", error.message);
    process.exit(1);
  });
