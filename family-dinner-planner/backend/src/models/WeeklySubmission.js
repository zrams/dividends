const mongoose = require("mongoose");

const weeklySubmissionSchema = new mongoose.Schema(
  {
    userId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "User",
      required: true
    },
    weekStartDate: {
      type: Date,
      required: true
    },
    choices: {
      type: [
        {
          type: mongoose.Schema.Types.ObjectId,
          ref: "Dinner",
          required: true
        }
      ],
      validate: {
        validator(value) {
          return Array.isArray(value) && value.length >= 1 && value.length <= 2;
        },
        message: "Each weekly submission must include 1-2 dinner choices."
      }
    }
  },
  {
    timestamps: true
  }
);

weeklySubmissionSchema.index({ userId: 1, weekStartDate: 1 }, { unique: true });

module.exports = mongoose.model("WeeklySubmission", weeklySubmissionSchema);
