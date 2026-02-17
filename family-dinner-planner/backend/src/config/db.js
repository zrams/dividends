const mongoose = require("mongoose");

async function connectDb() {
  if (!process.env.MONGO_URI) {
    throw new Error("MONGO_URI is required in environment variables");
  }

  await mongoose.connect(process.env.MONGO_URI);
  console.log("MongoDB connected");
}

module.exports = connectDb;
