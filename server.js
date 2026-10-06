require("dotenv").config();

const express = require("express");
const cors = require("cors");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");

const app = express();

app.disable("x-powered-by");
app.use(helmet());
app.use(express.json({ limit: "100kb" }));

app.use(
  cors({
    origin: function (origin, callback) {
      if (
        !origin ||
        /^https:\/\/[^.]+\.officescripts\.microsoftusercontent\.com$/.test(
          origin,
        )
      ) {
        return callback(null, true);
      }

      return callback(new Error("Origin is not permitted"));
    },
    methods: ["GET", "POST", "PUT", "OPTIONS"],
    allowedHeaders: ["Content-Type", "X-Integration-Key"],
  }),
);

app.use(
  rateLimit({
    windowMs: 60 * 1000,
    max: 30,
  }),
);

app.get("/api/health", function (req, res) {
  res.json({
    success: true,
    service: "TechOps Jira API",
  });
});

app.post("/api/jira/issues", async function (req, res) {
  try {
    const {
      summary,
      description,
      service,
      taskSource,
      taskDate,
      status,
      resolution,
      resolutionBy,
    } = req.body;

    if (!summary || typeof summary !== "string") {
      return res.status(400).json({
        success: false,
        message: "Summary is required.",
      });
    }

    if (summary.length > 255) {
      return res.status(400).json({
        success: false,
        message: "Summary cannot exceed 255 characters.",
      });
    }

    const jiraDescription = [
      description ? `Task Description:\n${description}` : "",
      service ? `Service: ${service}` : "",
      taskSource ? `Task Source: ${taskSource}` : "",
      taskDate ? `Task Date: ${taskDate}` : "",
      status ? `Excel Status: ${status}` : "",
      resolution ? `Resolution:\n${resolution}` : "",
      resolutionBy ? `Resolution By: ${resolutionBy}` : "",
    ]
      .filter(Boolean)
      .join("\n\n");

    const labels = ["excel-techops"];

    if (service) {
      labels.push(
        `service-${service}`
          .toLowerCase()
          .replace(/[^a-z0-9-]+/g, "-")
          .replace(/^-+|-+$/g, ""),
      );
    }

    const jiraPayload = {
      fields: {
        project: {
          key: process.env.JIRA_PROJECT_KEY,
        },
        issuetype: {
          name: process.env.JIRA_ISSUE_TYPE,
        },
        summary: summary.trim(),
        description: jiraDescription,
        priority: {
          name: "Medium",
        },
        labels,
      },
    };

    const jiraResponse = await fetch(
      `${process.env.JIRA_BASE_URL}/rest/api/2/issue`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${process.env.JIRA_TOKEN}`,
          Accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify(jiraPayload),
      },
    );

    const responseText = await jiraResponse.text();

    let jiraResult = {};

    if (responseText) {
      try {
        jiraResult = JSON.parse(responseText);
      } catch {
        jiraResult = {
          rawResponse: responseText,
        };
      }
    }

    if (!jiraResponse.ok) {
      console.error("Jira request failed", {
        status: jiraResponse.status,
        response: jiraResult,
      });

      return res.status(jiraResponse.status).json({
        success: false,
        message: "Jira rejected the issue request.",
        jiraStatus: jiraResponse.status,
        jiraErrors: jiraResult.errors || jiraResult.errorMessages || [],
      });
    }

    return res.status(201).json({
      success: true,
      jiraKey: jiraResult.key,
      jiraUrl: `${process.env.JIRA_BASE_URL}/browse/${jiraResult.key}`,
      message: "Jira issue created successfully.",
    });
  } catch (error) {
    console.error("Unexpected Jira integration error", error);

    return res.status(500).json({
      success: false,
      message: "The Jira integration encountered an internal error.",
    });
  }
});

app.use(function (error, req, res, next) {
  console.error(error);

  res.status(500).json({
    success: false,
    message: "Unexpected server error.",
  });
});

const port = Number(process.env.PORT || 3000);

app.listen(port, function () {
  console.log(`TechOps Jira API listening on port ${port}`);
});
``;
