require("dotenv").config();

const express = require("express");
const cors = require("cors");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");

const app = express();

/*
|--------------------------------------------------------------------------
| Configuration checks
|--------------------------------------------------------------------------
*/

const requiredEnvironmentVariables = [
  "JIRA_BASE_URL",
  "JIRA_PROJECT_KEY",
  "JIRA_ISSUE_TYPE",
  "JIRA_TOKEN",
];

const missingEnvironmentVariables = requiredEnvironmentVariables.filter(
  function (variableName) {
    return !process.env[variableName];
  },
);

if (missingEnvironmentVariables.length > 0) {
  console.error(
    `Missing environment variables: ${missingEnvironmentVariables.join(", ")}`,
  );

  process.exit(1);
}

/*
|--------------------------------------------------------------------------
| Middleware
|--------------------------------------------------------------------------
*/

app.disable("x-powered-by");

app.use(helmet());

app.use(
  express.json({
    limit: "100kb",
  }),
);

app.use(
  cors({
    origin: function (origin, callback) {
      /*
       * PowerShell and other server-side clients may not send Origin.
       * Office Scripts uses a Microsoft-hosted origin.
       */
      if (
        !origin ||
        /^https:\/\/[^.]+\.officescripts\.microsoftusercontent\.com$/.test(
          origin,
        )
      ) {
        return callback(null, true);
      }

      return callback(new Error("Origin is not permitted."));
    },

    methods: ["GET", "POST", "PUT", "OPTIONS"],

    allowedHeaders: ["Content-Type", "Accept", "X-Integration-Key"],
  }),
);

app.use(
  rateLimit({
    windowMs: 60 * 1000,
    max: 30,
    standardHeaders: true,
    legacyHeaders: false,
  }),
);

/*
|--------------------------------------------------------------------------
| General helper functions
|--------------------------------------------------------------------------
*/

function normalizeText(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ");
}

function normalizeDate(value) {
  const input = String(value || "").trim();

  if (!input) {
    return "";
  }

  /*
   * Jira-compatible YYYY-MM-DD format.
   */
  if (/^\d{4}-\d{2}-\d{2}$/.test(input)) {
    const parts = input.split("-");
    const year = Number(parts[0]);
    const month = Number(parts[1]);
    const day = Number(parts[2]);

    const date = new Date(Date.UTC(year, month - 1, day));

    if (
      date.getUTCFullYear() !== year ||
      date.getUTCMonth() + 1 !== month ||
      date.getUTCDate() !== day
    ) {
      throw new Error(`Invalid date "${input}".`);
    }

    return input;
  }

  /*
   * Excel workbook format:
   * MM/DD/YYYY
   *
   * Examples:
   * 10/06/2026
   * 10/6/2026
   */
  const slashDateMatch = input.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (slashDateMatch) {
    const monthNumber = Number(slashDateMatch[1]);
    const dayNumber = Number(slashDateMatch[2]);
    const yearNumber = Number(slashDateMatch[3]);

    const date = new Date(Date.UTC(yearNumber, monthNumber - 1, dayNumber));

    if (
      date.getUTCFullYear() !== yearNumber ||
      date.getUTCMonth() + 1 !== monthNumber ||
      date.getUTCDate() !== dayNumber
    ) {
      throw new Error(`Invalid date "${input}".`);
    }

    const month = String(monthNumber).padStart(2, "0");
    const day = String(dayNumber).padStart(2, "0");

    return `${yearNumber}-${month}-${day}`;
  }

  throw new Error(`Invalid date "${input}". Use MM/DD/YYYY or YYYY-MM-DD.`);
}

function createServiceLabel(service) {
  const normalizedService = String(service || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9-]+/g, "-")
    .replace(/^-+|-+$/g, "");

  return normalizedService ? `service-${normalizedService}` : "";
}

async function parseResponse(response) {
  const responseText = await response.text();

  if (!responseText) {
    return {};
  }

  try {
    return JSON.parse(responseText);
  } catch {
    return {
      rawResponse: responseText,
    };
  }
}

/*
|--------------------------------------------------------------------------
| Jira assignable-user lookup
|--------------------------------------------------------------------------
*/

async function resolveAssignableUser(userInput) {
  const requestedUser = String(userInput || "").trim();

  if (
    !requestedUser ||
    normalizeText(requestedUser) === "unassigned" ||
    normalizeText(requestedUser) === "automatic"
  ) {
    return null;
  }

  const searchUrl =
    `${process.env.JIRA_BASE_URL}` +
    `/rest/api/2/user/assignable/search` +
    `?project=${encodeURIComponent(process.env.JIRA_PROJECT_KEY)}` +
    `&username=${encodeURIComponent(requestedUser)}` +
    `&maxResults=100`;

  const response = await fetch(searchUrl, {
    method: "GET",

    headers: {
      Authorization: `Bearer ${process.env.JIRA_TOKEN}`,
      Accept: "application/json",
    },
  });

  const jiraResult = await parseResponse(response);

  if (!response.ok) {
    console.error("Jira assignee search failed", {
      status: response.status,
      response: jiraResult,
    });

    throw new Error(
      `Jira assignee search failed with HTTP ${response.status}.`,
    );
  }

  if (!Array.isArray(jiraResult)) {
    throw new Error(
      "Jira returned an unexpected response while searching for the assignee.",
    );
  }

  const normalizedInput = normalizeText(requestedUser);

  const exactMatches = jiraResult.filter(function (user) {
    if (user.active !== true) {
      return false;
    }

    const username = normalizeText(user.name);
    const displayName = normalizeText(user.displayName);
    const emailAddress = normalizeText(user.emailAddress);

    return (
      username === normalizedInput ||
      displayName === normalizedInput ||
      emailAddress === normalizedInput
    );
  });

  if (exactMatches.length === 1) {
    return exactMatches[0];
  }

  if (exactMatches.length > 1) {
    throw new Error(
      `More than one Jira account exactly matched "${requestedUser}". ` +
        "Enter the exact Jira username in Excel.",
    );
  }

  /*
   * Jira may return one partial match even if it was not an exact match.
   */
  const activeUsers = jiraResult.filter(function (user) {
    return user.active === true && user.name;
  });

  if (activeUsers.length === 1) {
    return activeUsers[0];
  }

  if (activeUsers.length > 1) {
    throw new Error(
      `Multiple Jira users matched "${requestedUser}". ` +
        "Enter the exact Jira username shown in the Assignee dropdown.",
    );
  }

  return null;
}

/*
|--------------------------------------------------------------------------
| Jira workflow functions
|--------------------------------------------------------------------------
*/

async function getIssueTransitions(issueKey) {
  const transitionsUrl =
    `${process.env.JIRA_BASE_URL}` +
    `/rest/api/2/issue/${encodeURIComponent(issueKey)}` +
    `/transitions?expand=transitions.fields`;

  const response = await fetch(transitionsUrl, {
    method: "GET",

    headers: {
      Authorization: `Bearer ${process.env.JIRA_TOKEN}`,
      Accept: "application/json",
    },
  });

  const jiraResult = await parseResponse(response);

  if (!response.ok) {
    console.error("Could not read Jira transitions", {
      issueKey,
      status: response.status,
      response: jiraResult,
    });

    throw new Error(
      `Could not retrieve workflow transitions for ${issueKey}. ` +
        `Jira returned HTTP ${response.status}.`,
    );
  }

  return Array.isArray(jiraResult.transitions) ? jiraResult.transitions : [];
}

function getTransitionDetails(transition) {
  return {
    id: String(transition.id || ""),
    name: String(transition.name || ""),
    destination:
      transition.to && transition.to.name ? String(transition.to.name) : "",
  };
}

function transitionMatches(transition, possibleNames) {
  const transitionName = normalizeText(transition.name);

  const destinationName = normalizeText(
    transition.to && transition.to.name ? transition.to.name : "",
  );

  return possibleNames.some(function (possibleName) {
    const normalizedPossibleName = normalizeText(possibleName);

    return (
      transitionName === normalizedPossibleName ||
      destinationName === normalizedPossibleName
    );
  });
}

async function performTransition(issueKey, possibleNames) {
  const transitions = await getIssueTransitions(issueKey);

  const availableTransitions = transitions.map(getTransitionDetails);

  const matchingTransition = transitions.find(function (transition) {
    return transitionMatches(transition, possibleNames);
  });

  if (!matchingTransition) {
    return {
      success: false,
      changed: false,
      message:
        `No available Jira transition matched: ` +
        `${possibleNames.join(", ")}.`,
      availableTransitions,
    };
  }

  const transitionPayload = {
    transition: {
      id: String(matchingTransition.id),
    },
  };

  const transitionUrl =
    `${process.env.JIRA_BASE_URL}` +
    `/rest/api/2/issue/${encodeURIComponent(issueKey)}` +
    `/transitions`;

  const response = await fetch(transitionUrl, {
    method: "POST",

    headers: {
      Authorization: `Bearer ${process.env.JIRA_TOKEN}`,
      Accept: "application/json",
      "Content-Type": "application/json",
    },

    body: JSON.stringify(transitionPayload),
  });

  const jiraResult = await parseResponse(response);

  if (!response.ok) {
    console.error("Jira transition failed", {
      issueKey,
      transitionId: matchingTransition.id,
      transitionName: matchingTransition.name,
      status: response.status,
      response: jiraResult,
    });

    return {
      success: false,
      changed: false,
      transitionId: String(matchingTransition.id),
      transitionName: matchingTransition.name,
      message:
        `Jira transition "${matchingTransition.name}" failed ` +
        `with HTTP ${response.status}.`,
      jiraStatus: response.status,
      jiraErrors: jiraResult.errors || jiraResult.errorMessages || jiraResult,
      availableTransitions,
    };
  }

  const transitionDetails = getTransitionDetails(matchingTransition);

  return {
    success: true,
    changed: true,
    transitionId: transitionDetails.id,
    transitionName: transitionDetails.name,
    destinationStatus: transitionDetails.destination,
    message:
      `${issueKey} moved to ` +
      `"${transitionDetails.destination || transitionDetails.name}".`,
    availableTransitions,
  };
}

/*
|--------------------------------------------------------------------------
| Determine Jira status from End Date
|--------------------------------------------------------------------------
|
| End Date blank:
|   Initial status is To Do.
|   Express moves the issue to In Progress.
|
| End Date present:
|   Initial status is To Do.
|   Express moves the issue directly to Done.
|
*/

async function applyDateBasedStatus(issueKey, endDate) {
  const hasEndDate = Boolean(endDate && String(endDate).trim());

  const inProgressTransitionNames = [
    "In Progress",
    "Dev - In Progress",
    "Start Progress",
    "Start Work",
  ];

  const doneTransitionNames = ["Done"];

  /*
   * No End Date means the task is still active.
   */
  if (!hasEndDate) {
    const inProgressResult = await performTransition(
      issueKey,
      inProgressTransitionNames,
    );

    return {
      requestedStatus: "In Progress",
      statusChanged: inProgressResult.success,
      actualResult: inProgressResult,
      message: inProgressResult.success
        ? `${issueKey} moved to In Progress because End Date is blank.`
        : `${issueKey} was created, but the In Progress transition was not available.`,
    };
  }

  /*
   * End Date exists, so the task is complete.
   *
   * The Jira workflow shown allows:
   * To Do -> Done
   */
  const doneResult = await performTransition(issueKey, doneTransitionNames);

  return {
    requestedStatus: "Done",
    statusChanged: doneResult.success,
    actualResult: doneResult,
    message: doneResult.success
      ? `${issueKey} moved to Done because End Date is present.`
      : `${issueKey} was created, but the Done transition was not available.`,
  };
}

/*
|--------------------------------------------------------------------------
| Get assignable Jira users
|--------------------------------------------------------------------------
|
| Examples:
|
| GET /api/jira/assignees?search=Faijul
| GET /api/jira/assignees?search=faijul.islam
|--------------------------------------------------------------------------
*/

app.get("/api/jira/assignees", async function (req, res) {
  try {
    const search = String(req.query.search || "").trim();

    const searchUrl =
      `${process.env.JIRA_BASE_URL}` +
      `/rest/api/2/user/assignable/search` +
      `?project=${encodeURIComponent(process.env.JIRA_PROJECT_KEY)}` +
      `&username=${encodeURIComponent(search)}` +
      `&maxResults=100`;

    const response = await fetch(searchUrl, {
      method: "GET",

      headers: {
        Authorization: `Bearer ${process.env.JIRA_TOKEN}`,
        Accept: "application/json",
      },
    });

    const jiraResult = await parseResponse(response);

    if (!response.ok) {
      return res.status(response.status).json({
        success: false,
        message: "Jira rejected the assignable-user search.",
        jiraStatus: response.status,
        jiraErrors: jiraResult.errors || jiraResult.errorMessages || jiraResult,
      });
    }

    const users = Array.isArray(jiraResult)
      ? jiraResult
          .filter(function (user) {
            return user.active === true;
          })
          .map(function (user) {
            return {
              username: user.name || "",
              displayName: user.displayName || "",
              emailAddress: user.emailAddress || "",
              active: user.active === true,
            };
          })
      : [];

    return res.json({
      success: true,
      projectKey: process.env.JIRA_PROJECT_KEY,
      count: users.length,
      users,
    });
  } catch (error) {
    console.error("Assignable-user endpoint failed", error);

    return res.status(500).json({
      success: false,
      message:
        error instanceof Error
          ? error.message
          : "Could not retrieve Jira assignees.",
    });
  }
});

/*
|--------------------------------------------------------------------------
| Health endpoint
|--------------------------------------------------------------------------
*/

app.get("/api/health", function (req, res) {
  res.json({
    success: true,
    service: "TechOps Jira API",
    projectKey: process.env.JIRA_PROJECT_KEY,
    issueType: process.env.JIRA_ISSUE_TYPE,
  });
});

/*
|--------------------------------------------------------------------------
| Create Jira issue
|--------------------------------------------------------------------------
*/

app.post("/api/jira/issues", async function (req, res) {
  try {
    const {
      summary,
      description,
      service,
      taskSource,
      taskDate,
      resolution,
      resolutionBy,
      assignedTo,
      startDate,
      endDate,
    } = req.body;

    /*
     * Validate summary.
     */
    if (!summary || typeof summary !== "string") {
      return res.status(400).json({
        success: false,
        message: "Summary is required.",
      });
    }

    const cleanSummary = summary.trim();

    if (!cleanSummary) {
      return res.status(400).json({
        success: false,
        message: "Summary cannot be empty.",
      });
    }

    if (cleanSummary.length > 255) {
      return res.status(400).json({
        success: false,
        message: "Summary cannot exceed 255 characters.",
      });
    }

    /*
     * Normalize Excel dates.
     */
    let cleanTaskDate = "";
    let cleanStartDate = "";
    let cleanEndDate = "";

    try {
      cleanTaskDate = normalizeDate(taskDate);
      cleanStartDate = normalizeDate(startDate);
      cleanEndDate = normalizeDate(endDate);
    } catch (dateError) {
      return res.status(400).json({
        success: false,
        message:
          dateError instanceof Error
            ? dateError.message
            : "One or more dates are invalid.",
      });
    }

    /*
     * If Start Date was not provided separately,
     * use Task Date as Start Date.
     */
    if (!cleanStartDate && cleanTaskDate) {
      cleanStartDate = cleanTaskDate;
    }

    /*
     * Determine target Jira status.
     */
    const calculatedStatus = cleanEndDate ? "Done" : "In Progress";

    /*
     * Validate the requested Jira assignee.
     */
    let assignee = null;

    try {
      assignee = await resolveAssignableUser(assignedTo);
    } catch (assigneeError) {
      return res.status(400).json({
        success: false,
        message:
          assigneeError instanceof Error
            ? assigneeError.message
            : "Could not validate the requested Jira assignee.",
      });
    }

    const normalizedRequestedAssignee = normalizeText(assignedTo);

    if (
      assignedTo &&
      normalizedRequestedAssignee !== "unassigned" &&
      normalizedRequestedAssignee !== "automatic" &&
      !assignee
    ) {
      return res.status(400).json({
        success: false,
        message:
          `No active assignable Jira account matched ` +
          `"${assignedTo}" in project ` +
          `${process.env.JIRA_PROJECT_KEY}. ` +
          "Use the exact Jira username shown in the Assignee dropdown.",
      });
    }

    /*
     * Build Jira description.
     */
    const jiraDescription = [
      description ? `Task Description:\n${description}` : "",

      service ? `Service: ${service}` : "",

      taskSource ? `Task Source: ${taskSource}` : "",

      cleanTaskDate ? `Task Date: ${cleanTaskDate}` : "",

      cleanStartDate ? `Start Date: ${cleanStartDate}` : "",

      cleanEndDate ? `End Date: ${cleanEndDate}` : "",

      `Calculated Jira Status: ${calculatedStatus}`,

      resolution ? `Resolution:\n${resolution}` : "",

      resolutionBy ? `Resolution By: ${resolutionBy}` : "",

      assignee
        ? `Requested Assignee: ` +
          `${assignee.displayName || assignee.name} ` +
          `(${assignee.name})`
        : "Requested Assignee: Unassigned",
    ]
      .filter(Boolean)
      .join("\n\n");

    /*
     * Build labels.
     */
    const labels = ["excel-techops"];

    const serviceLabel = createServiceLabel(service);

    if (serviceLabel) {
      labels.push(serviceLabel);
    }

    if (cleanEndDate) {
      labels.push("excel-completed");
    } else {
      labels.push("excel-in-progress");
    }

    /*
     * Build Jira fields.
     */
    const jiraFields = {
      project: {
        key: process.env.JIRA_PROJECT_KEY,
      },

      issuetype: {
        name: process.env.JIRA_ISSUE_TYPE,
      },

      summary: cleanSummary,

      description: jiraDescription,

      priority: {
        name: "Medium",
      },

      labels,
    };

    /*
     * Add the verified assignee.
     */
    if (assignee) {
      jiraFields.assignee = {
        name: assignee.name,
      };
    }

    /*
     * Optional Jira custom Start Date field.
     */
    if (cleanStartDate && process.env.JIRA_START_DATE_FIELD) {
      jiraFields[process.env.JIRA_START_DATE_FIELD] = cleanStartDate;
    }

    /*
     * Optional Jira custom End Date field.
     */
    if (cleanEndDate && process.env.JIRA_END_DATE_FIELD) {
      jiraFields[process.env.JIRA_END_DATE_FIELD] = cleanEndDate;
    }

    /*
     * Optionally use End Date as Jira Due Date.
     */
    if (
      cleanEndDate &&
      normalizeText(process.env.USE_END_DATE_AS_DUE_DATE) === "true"
    ) {
      jiraFields.duedate = cleanEndDate;
    }

    const jiraPayload = {
      fields: jiraFields,
    };

    /*
     * Create Jira issue.
     */
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

    const jiraResult = await parseResponse(jiraResponse);

    if (!jiraResponse.ok) {
      console.error("Jira issue creation failed", {
        status: jiraResponse.status,
        response: jiraResult,
      });

      return res.status(jiraResponse.status).json({
        success: false,
        message: "Jira rejected the issue creation request.",
        jiraStatus: jiraResponse.status,
        jiraErrors: jiraResult.errors || jiraResult.errorMessages || jiraResult,
      });
    }

    const createdIssueKey = jiraResult.key;

    /*
     * Move the issue according to End Date:
     *
     * End Date blank   -> In Progress
     * End Date present -> Done
     */
    let statusResult;

    try {
      statusResult = await applyDateBasedStatus(createdIssueKey, cleanEndDate);
    } catch (statusError) {
      console.error("Issue created, but status update failed", {
        issueKey: createdIssueKey,
        calculatedStatus,
        error: statusError instanceof Error ? statusError.message : statusError,
      });

      statusResult = {
        requestedStatus: calculatedStatus,
        statusChanged: false,
        actualResult: {
          success: false,
          changed: false,
        },
        message:
          statusError instanceof Error
            ? statusError.message
            : "The Jira issue was created, but the " +
              "workflow status could not be updated.",
      };
    }

    /*
     * Return Jira issue details to Excel.
     *
     * A failed transition does not delete the created issue.
     * The response shows statusChanged=false.
     */
    return res.status(201).json({
      success: true,

      jiraKey: createdIssueKey,

      jiraUrl: `${process.env.JIRA_BASE_URL}` + `/browse/${createdIssueKey}`,

      assignee: assignee
        ? {
            username: assignee.name,
            displayName: assignee.displayName || assignee.name,
          }
        : null,

      taskDate: cleanTaskDate,

      startDate: cleanStartDate,

      endDate: cleanEndDate,

      calculatedStatus,

      statusChanged: Boolean(statusResult.statusChanged),

      statusMessage:
        statusResult.message || "No workflow status message was returned.",

      statusDetails: statusResult,

      message: statusResult.statusChanged
        ? `Jira issue ${createdIssueKey} was created ` +
          `and moved to ${calculatedStatus}.`
        : `Jira issue ${createdIssueKey} was created, ` +
          `but it could not be moved to ` +
          `${calculatedStatus}.`,
    });
  } catch (error) {
    console.error("Unexpected Jira integration error", error);

    return res.status(500).json({
      success: false,
      message:
        error instanceof Error
          ? error.message
          : "The Jira integration encountered " + "an internal error.",
    });
  }
});

/*
|--------------------------------------------------------------------------
| Express error handler
|--------------------------------------------------------------------------
*/

app.use(function (error, req, res, next) {
  console.error("Express middleware error", error);

  if (error instanceof Error && error.message === "Origin is not permitted.") {
    return res.status(403).json({
      success: false,
      message: "Request origin is not permitted.",
    });
  }

  return res.status(500).json({
    success: false,
    message: "Unexpected server error.",
  });
});

/*
|--------------------------------------------------------------------------
| Start server
|--------------------------------------------------------------------------
*/

const port = Number(process.env.PORT || 3000);

app.listen(port, function () {
  console.log(`TechOps Jira API listening on port ${port}`);

  console.log(`Jira project: ${process.env.JIRA_PROJECT_KEY}`);

  console.log(`Jira issue type: ${process.env.JIRA_ISSUE_TYPE}`);

  console.log("Date-based workflow: End Date blank = In Progress");

  console.log("Date-based workflow: End Date present = Done");
});
