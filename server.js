require("dotenv").config();

const express = require("express");
const cors = require("cors");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");

const app = express();

/*
|--------------------------------------------------------------------------
| Environment configuration
|--------------------------------------------------------------------------
*/

const requiredEnvironmentVariables = [
  "JIRA_BASE_URL",
  "JIRA_PROJECT_KEY",
  "JIRA_ISSUE_TYPE",
  "UPSTASH_REDIS_REST_URL",
  "UPSTASH_REDIS_REST_TOKEN",
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

const jiraBaseUrl = String(process.env.JIRA_BASE_URL).replace(/\/+$/, "");

const jiraProjectKey = String(process.env.JIRA_PROJECT_KEY)
  .trim()
  .toUpperCase();

const jiraIssueType = String(process.env.JIRA_ISSUE_TYPE).trim();

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

const allowedOriginPatterns = [
  /^https:\/\/[^.]+\.officescripts\.microsoftusercontent\.com$/i,
  /^https:\/\/[^.]+\.officeapps\.live\.com$/i,
  /^https:\/\/excel\.officeapps\.live\.com$/i,
  /^https:\/\/www\.office\.com$/i,
  /^https:\/\/www\.microsoft365\.com$/i,
];

app.use(
  cors({
    origin: function (origin, callback) {
      /*
       * Allow server-side requests that do not include Origin.
       */
      if (!origin) {
        return callback(null, true);
      }

      let isAllowedOrigin = false;

      for (let index = 0; index < allowedOriginPatterns.length; index++) {
        if (allowedOriginPatterns[index].test(origin)) {
          isAllowedOrigin = true;
          break;
        }
      }

      if (isAllowedOrigin) {
        return callback(null, true);
      }

      console.warn(`Blocked CORS origin: ${origin}`);

      return callback(new Error("Origin is not permitted."));
    },

    methods: ["GET", "POST", "PUT", "OPTIONS"],

    allowedHeaders: ["Content-Type", "Accept", "X-Integration-Key"],

    exposedHeaders: ["Content-Type"],

    credentials: false,

    optionsSuccessStatus: 204,

    maxAge: 86400,
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
| Optional Express API authentication
|--------------------------------------------------------------------------
|
| If EXPRESS_INTEGRATION_KEY is present in .env, every Jira API request
| must include:
|
| X-Integration-Key: your-secret-key
|
| The health endpoint remains available without the key.
|--------------------------------------------------------------------------
*/

function validateIntegrationKey(req, res, next) {
  const requiredKey = String(process.env.EXPRESS_INTEGRATION_KEY || "").trim();

  /*
   * If no key is configured, allow access.
   * This is useful for local testing.
   */
  if (!requiredKey) {
    return next();
  }

  const suppliedKey = String(req.get("X-Integration-Key") || "").trim();

  if (!suppliedKey || suppliedKey !== requiredKey) {
    return res.status(401).json({
      success: false,
      message: "The Express integration key is missing or invalid.",
    });
  }

  return next();
}

app.use("/api/jira", validateIntegrationKey);

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

  throw new Error(
    `Invalid date "${input}". ` + "Use MM/DD/YYYY or YYYY-MM-DD.",
  );
}

function createServiceLabel(service) {
  const normalizedService = String(service || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9-]+/g, "-")
    .replace(/^-+|-+$/g, "");

  return normalizedService ? `service-${normalizedService}` : "";
}

function createIntegrationLabel(integrationId) {
  const cleanId = String(integrationId || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9-]+/g, "-")
    .replace(/^-+|-+$/g, "");

  return cleanId ? `excel-id-${cleanId}` : "";
}

function isUnassignedValue(value) {
  const normalizedValue = normalizeText(value);

  return (
    !normalizedValue ||
    normalizedValue === "unassigned" ||
    normalizedValue === "automatic"
  );
}

function validateJiraKey(value) {
  const jiraKey = String(value || "")
    .trim()
    .toUpperCase();

  const pattern = new RegExp(
    `^${jiraProjectKey.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}-\\d+$`,
  );

  return pattern.test(jiraKey);
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

function getJiraError(result) {
  if (!result) {
    return result;
  }

  return result.errors || result.errorMessages || result;
}

/*
|--------------------------------------------------------------------------
| Upstash Jira token store
|--------------------------------------------------------------------------
*/
const jiraTokenStoreKey = String(
  process.env.JIRA_TOKEN_STORE_KEY || "jira-user-tokens",
).trim();
const tokenCacheTtlMs = Number(process.env.JIRA_TOKEN_CACHE_TTL_MS || 60000);
let jiraTokenCache = { loadedAt: 0, entries: [] };

function parseStoredTokenDocument(value) {
  let current = value;
  for (let attempt = 0; attempt < 2 && typeof current === "string"; attempt++) {
    const text = current.trim();
    if (!text) return { data: { tokens: [] } };
    try {
      current = JSON.parse(text);
    } catch {
      throw new Error(
        `Upstash key "${jiraTokenStoreKey}" does not contain valid JSON.`,
      );
    }
  }
  return current;
}

function extractTokenEntries(document) {
  const tokens =
    document && document.data && Array.isArray(document.data.tokens)
      ? document.data.tokens
      : [];
  const entries = tokens
    .map(function (entry) {
      return {
        name: String(entry && entry.name ? entry.name : "").trim(),
        token: String(entry && entry.token ? entry.token : "").trim(),
      };
    })
    .filter(function (entry) {
      return Boolean(entry.name && entry.token);
    });
  const seen = new Set();
  const duplicates = new Set();
  entries.forEach(function (entry) {
    const name = normalizeText(entry.name);
    if (seen.has(name)) duplicates.add(entry.name);
    seen.add(name);
  });
  if (duplicates.size > 0) {
    throw new Error(
      `Upstash contains duplicate Jira token names: ${Array.from(duplicates).join(", ")}.`,
    );
  }
  return entries;
}

async function loadJiraTokenEntries(forceRefresh = false) {
  const now = Date.now();
  if (
    !forceRefresh &&
    jiraTokenCache.entries.length > 0 &&
    now - jiraTokenCache.loadedAt < tokenCacheTtlMs
  ) {
    return jiraTokenCache.entries;
  }
  const url = String(process.env.UPSTASH_REDIS_REST_URL).replace(/\/+$/, "");
  const response = await fetch(
    `${url}/get/${encodeURIComponent(jiraTokenStoreKey)}`,
    {
      method: "GET",
      headers: {
        Authorization: `Bearer ${process.env.UPSTASH_REDIS_REST_TOKEN}`,
        Accept: "application/json",
      },
    },
  );
  const result = await parseResponse(response);
  if (!response.ok)
    throw new Error(
      `Upstash token lookup failed with HTTP ${response.status}.`,
    );
  if (result.result === null || result.result === undefined) {
    throw new Error(`Upstash key "${jiraTokenStoreKey}" was not found.`);
  }
  const entries = extractTokenEntries(parseStoredTokenDocument(result.result));
  jiraTokenCache = { loadedAt: now, entries };
  return entries;
}

async function resolveJiraTokenForAssignee(assignedTo) {
  if (!isUnassignedValue(assignedTo)) {
    const requested = normalizeText(assignedTo);
    const entries = await loadJiraTokenEntries();
    const match = entries.find(function (entry) {
      return normalizeText(entry.name) === requested;
    });
    if (match)
      return {
        token: match.token,
        source: "upstash-assignee-token",
        user: match.name,
      };
  }
  return {
    token: String(process.env.JIRA_TOKEN).trim(),
    source: "environment-default",
    user: "",
  };
}

/*
|--------------------------------------------------------------------------
| Jira assignee lookup
|--------------------------------------------------------------------------
*/

async function resolveAssignableUser(userInput, jiraToken) {
  const requestedUser = String(userInput || "").trim();

  if (isUnassignedValue(requestedUser)) {
    return null;
  }

  const searchUrl =
    `${jiraBaseUrl}` +
    "/rest/api/2/user/assignable/search" +
    `?project=${encodeURIComponent(jiraProjectKey)}` +
    `&username=${encodeURIComponent(requestedUser)}` +
    "&maxResults=100";

  const response = await fetch(searchUrl, {
    method: "GET",

    headers: {
      Authorization: `Bearer ${jiraToken}`,
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
      `Jira assignee search failed with ` + `HTTP ${response.status}.`,
    );
  }

  if (!Array.isArray(jiraResult)) {
    throw new Error(
      "Jira returned an unexpected response " +
        "while searching for the assignee.",
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
      `More than one Jira account exactly matched ` +
        `"${requestedUser}". Use the exact Jira username.`,
    );
  }

  const activeUsers = jiraResult.filter(function (user) {
    return user.active === true && Boolean(user.name);
  });

  if (activeUsers.length === 1) {
    return activeUsers[0];
  }

  if (activeUsers.length > 1) {
    throw new Error(
      `Multiple Jira users matched "${requestedUser}". ` +
        "Use the exact Jira username shown in Jira.",
    );
  }

  return null;
}

/*
|--------------------------------------------------------------------------
| Duplicate detection
|--------------------------------------------------------------------------
*/

async function findIssueByIntegrationId(integrationId, jiraToken) {
  const integrationLabel = createIntegrationLabel(integrationId);

  if (!integrationLabel) {
    return null;
  }

  const jql =
    `project = ${jiraProjectKey} ` +
    `AND labels = "${integrationLabel}" ` +
    "ORDER BY created DESC";

  const searchUrl =
    `${jiraBaseUrl}/rest/api/2/search` +
    `?jql=${encodeURIComponent(jql)}` +
    "&maxResults=2" +
    "&fields=key,summary,status";

  const response = await fetch(searchUrl, {
    method: "GET",

    headers: {
      Authorization: `Bearer ${jiraToken}`,
      Accept: "application/json",
    },
  });

  const result = await parseResponse(response);

  if (!response.ok) {
    console.error("Jira duplicate search failed", {
      status: response.status,
      response: result,
    });

    throw new Error(
      `Jira duplicate search failed with ` + `HTTP ${response.status}.`,
    );
  }

  const issues = Array.isArray(result.issues) ? result.issues : [];

  if (issues.length > 1) {
    throw new Error(
      `Multiple Jira issues use Integration ID ` +
        `"${integrationId}". Resolve the duplicate ` +
        "labels in Jira before continuing.",
    );
  }

  return issues.length === 1 ? issues[0] : null;
}

/*
|--------------------------------------------------------------------------
| Jira issue information
|--------------------------------------------------------------------------
*/

async function getJiraIssue(issueKey, jiraToken) {
  const response = await fetch(
    `${jiraBaseUrl}/rest/api/2/issue/` +
      `${encodeURIComponent(issueKey)}` +
      "?fields=key,summary,status,resolution,labels",
    {
      method: "GET",

      headers: {
        Authorization: `Bearer ${jiraToken}`,
        Accept: "application/json",
      },
    },
  );

  const result = await parseResponse(response);

  if (response.status === 404) {
    return null;
  }

  if (!response.ok) {
    throw new Error(
      `Could not read Jira issue ${issueKey}. ` +
        `Jira returned HTTP ${response.status}.`,
    );
  }

  return result;
}

async function getIssueStatus(issueKey, jiraToken) {
  const issue = await getJiraIssue(issueKey, jiraToken);

  if (!issue) {
    throw new Error(`Jira issue ${issueKey} was not found.`);
  }

  return {
    status:
      issue.fields && issue.fields.status && issue.fields.status.name
        ? String(issue.fields.status.name)
        : "",

    resolution:
      issue.fields && issue.fields.resolution && issue.fields.resolution.name
        ? String(issue.fields.resolution.name)
        : "",
  };
}

/*
|--------------------------------------------------------------------------
| Jira workflow transitions
|--------------------------------------------------------------------------
*/

async function getIssueTransitions(issueKey, jiraToken) {
  const transitionsUrl =
    `${jiraBaseUrl}/rest/api/2/issue/` +
    `${encodeURIComponent(issueKey)}` +
    "/transitions?expand=transitions.fields";

  const response = await fetch(transitionsUrl, {
    method: "GET",

    headers: {
      Authorization: `Bearer ${jiraToken}`,
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
      `Could not retrieve workflow transitions ` +
        `for ${issueKey}. Jira returned ` +
        `HTTP ${response.status}.`,
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

async function performTransition(issueKey, possibleNames, jiraToken) {
  const transitions = await getIssueTransitions(issueKey, jiraToken);

  const availableTransitions = transitions.map(getTransitionDetails);

  const matchingTransition = transitions.find(function (transition) {
    return transitionMatches(transition, possibleNames);
  });

  if (!matchingTransition) {
    return {
      success: false,
      changed: false,

      message:
        "No available Jira transition matched: " +
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
    `${jiraBaseUrl}/rest/api/2/issue/` +
    `${encodeURIComponent(issueKey)}` +
    "/transitions";

  const response = await fetch(transitionUrl, {
    method: "POST",

    headers: {
      Authorization: `Bearer ${jiraToken}`,
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
        `Jira transition ` +
        `"${matchingTransition.name}" failed ` +
        `with HTTP ${response.status}.`,

      jiraStatus: response.status,

      jiraErrors: getJiraError(jiraResult),

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
| Date-based Jira status
|--------------------------------------------------------------------------
|
| Resolution Date blank:
|   Desired Jira status is In Progress.
|
| Resolution Date populated:
|   Desired Jira status is Done.
|--------------------------------------------------------------------------
*/

async function applyDateBasedStatus(issueKey, endDate, jiraToken) {
  const hasEndDate = Boolean(endDate && String(endDate).trim());

  const currentIssue = await getIssueStatus(issueKey, jiraToken);

  const currentStatus = normalizeText(currentIssue.status);

  const doneStatuses = ["done", "resolved", "closed", "completed"];

  const inProgressStatuses = ["in progress", "dev in progress"];

  const currentIsDone = doneStatuses.includes(currentStatus);

  const currentIsInProgress = inProgressStatuses.includes(currentStatus);

  /*
   * End Date exists and issue is already complete.
   */
  if (hasEndDate && currentIsDone) {
    return {
      requestedStatus: "Done",
      statusChanged: true,
      alreadyCorrect: true,

      actualResult: {
        success: true,
        changed: false,
        currentStatus: currentIssue.status,
        resolution: currentIssue.resolution,
      },

      message: `${issueKey} is already in ` + `${currentIssue.status}.`,
    };
  }

  /*
   * End Date is blank and issue is already In Progress.
   */
  if (!hasEndDate && currentIsInProgress) {
    return {
      requestedStatus: "In Progress",
      statusChanged: true,
      alreadyCorrect: true,

      actualResult: {
        success: true,
        changed: false,
        currentStatus: currentIssue.status,
      },

      message: `${issueKey} is already in ` + `${currentIssue.status}.`,
    };
  }

  /*
   * End Date exists, move to Done.
   */
  if (hasEndDate) {
    const doneResult = await performTransition(
      issueKey,
      [
        "Done",
        "Resolve",
        "Resolve Issue",
        "Resolved",
        "Complete",
        "Completed",
        "Close",
        "Close Issue",
      ],
      jiraToken,
    );

    return {
      requestedStatus: "Done",
      statusChanged: doneResult.success,
      actualResult: doneResult,

      message: doneResult.success
        ? `${issueKey} moved to Done because ` + "Resolution Date is present."
        : `${issueKey} could not be moved to Done. ` + doneResult.message,
    };
  }

  /*
   * End Date is blank.
   *
   * If the ticket is already Done, the workflow may expose a
   * Reopen transition whose destination is In Progress or To Do.
   */
  const inProgressResult = await performTransition(
    issueKey,
    [
      "In Progress",
      "Dev - In Progress",
      "Start Progress",
      "Start Work",
      "Reopen",
      "Re-open",
      "Reopen Issue",
      "Return to In Progress",
    ],
    jiraToken,
  );

  return {
    requestedStatus: "In Progress",
    statusChanged: inProgressResult.success,
    actualResult: inProgressResult,

    message: inProgressResult.success
      ? `${issueKey} moved to In Progress ` +
        "because Resolution Date is blank."
      : `${issueKey} could not be moved to ` +
        "In Progress. " +
        inProgressResult.message,
  };
}

/*
|--------------------------------------------------------------------------
| Prepare issue content
|--------------------------------------------------------------------------
*/

async function prepareIssueData(body) {
  const jiraCredential = await resolveJiraTokenForAssignee(body.assignedTo);
  const summary = body.summary;
  const description = body.description;
  const service = body.service;
  const taskSource = body.taskSource;
  const taskDate = body.taskDate;
  const resolution = body.resolution;
  const resolutionBy = body.resolutionBy;
  const assignedTo = body.assignedTo;
  const startDate = body.startDate;
  const endDate = body.endDate;
  const integrationId = body.integrationId;

  if (!summary || typeof summary !== "string") {
    throw new Error("Summary is required.");
  }

  const cleanSummary = summary.trim();

  if (!cleanSummary) {
    throw new Error("Summary cannot be empty.");
  }

  if (cleanSummary.length > 255) {
    throw new Error("Summary cannot exceed 255 characters.");
  }

  const cleanIntegrationId = String(integrationId || "").trim();

  if (!cleanIntegrationId) {
    throw new Error("Integration ID is required.");
  }

  let cleanTaskDate = "";
  let cleanStartDate = "";
  let cleanEndDate = "";

  cleanTaskDate = normalizeDate(taskDate);

  cleanStartDate = normalizeDate(startDate);

  cleanEndDate = normalizeDate(endDate);

  if (!cleanStartDate && cleanTaskDate) {
    cleanStartDate = cleanTaskDate;
  }

  const calculatedStatus = cleanEndDate ? "Done" : "In Progress";

  const assignee = await resolveAssignableUser(
    assignedTo,
    jiraCredential.token,
  );

  if (!isUnassignedValue(assignedTo) && !assignee) {
    throw new Error(
      `No active assignable Jira account ` +
        `matched "${assignedTo}" in project ` +
        `${jiraProjectKey}. Use the exact ` +
        "Jira username.",
    );
  }

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

    `Excel Integration ID: ${cleanIntegrationId}`,

    assignee
      ? "Requested Assignee: " +
        `${assignee.displayName || assignee.name} ` +
        `(${assignee.name})`
      : "Requested Assignee: Unassigned",
  ]
    .filter(Boolean)
    .join("\n\n");

  const labels = ["excel-techops"];

  const integrationLabel = createIntegrationLabel(cleanIntegrationId);

  if (integrationLabel) {
    labels.push(integrationLabel);
  }

  const serviceLabel = createServiceLabel(service);

  if (serviceLabel) {
    labels.push(serviceLabel);
  }

  labels.push(cleanEndDate ? "excel-completed" : "excel-in-progress");

  return {
    cleanSummary,
    cleanIntegrationId,
    cleanTaskDate,
    cleanStartDate,
    cleanEndDate,
    calculatedStatus,
    assignee,
    jiraDescription,
    labels,
    jiraToken: jiraCredential.token,
    jiraTokenSource: jiraCredential.source,
    jiraTokenUser: jiraCredential.user,
  };
}

/*
|--------------------------------------------------------------------------
| Build Jira fields
|--------------------------------------------------------------------------
*/

function buildCreateFields(prepared) {
  const jiraFields = {
    project: {
      key: jiraProjectKey,
    },

    issuetype: {
      name: jiraIssueType,
    },

    summary: prepared.cleanSummary,

    description: prepared.jiraDescription,

    priority: {
      name: process.env.JIRA_DEFAULT_PRIORITY || "Medium",
    },

    labels: prepared.labels,
  };

  if (prepared.assignee) {
    jiraFields.assignee = {
      name: prepared.assignee.name,
    };
  }

  if (prepared.cleanStartDate && process.env.JIRA_START_DATE_FIELD) {
    jiraFields[process.env.JIRA_START_DATE_FIELD] = prepared.cleanStartDate;
  }

  if (prepared.cleanEndDate && process.env.JIRA_END_DATE_FIELD) {
    jiraFields[process.env.JIRA_END_DATE_FIELD] = prepared.cleanEndDate;
  }

  if (
    prepared.cleanEndDate &&
    normalizeText(process.env.USE_END_DATE_AS_DUE_DATE) === "true"
  ) {
    jiraFields.duedate = prepared.cleanEndDate;
  }

  return jiraFields;
}

function buildUpdateFields(prepared) {
  const jiraFields = {
    summary: prepared.cleanSummary,

    description: prepared.jiraDescription,

    labels: prepared.labels,

    assignee: prepared.assignee
      ? {
          name: prepared.assignee.name,
        }
      : null,
  };

  if (process.env.JIRA_START_DATE_FIELD) {
    jiraFields[process.env.JIRA_START_DATE_FIELD] =
      prepared.cleanStartDate || null;
  }

  if (process.env.JIRA_END_DATE_FIELD) {
    jiraFields[process.env.JIRA_END_DATE_FIELD] = prepared.cleanEndDate || null;
  }

  if (normalizeText(process.env.USE_END_DATE_AS_DUE_DATE) === "true") {
    jiraFields.duedate = prepared.cleanEndDate || null;
  }

  return jiraFields;
}

/*
|--------------------------------------------------------------------------
| Health endpoint
|--------------------------------------------------------------------------
*/

app.get("/api/health", async function (req, res) {
  try {
    const tokenEntries = await loadJiraTokenEntries();
    return res.json({
      success: true,
      service: "TechOps Jira API",
      projectKey: jiraProjectKey,
      issueType: jiraIssueType,
      tokenStore: {
        provider: "Upstash Redis",
        key: jiraTokenStoreKey,
        availableTokenNames: tokenEntries.map(function (entry) {
          return entry.name;
        }),
        count: tokenEntries.length,
        fallbackConfigured: Boolean(
          String(process.env.JIRA_TOKEN || "").trim(),
        ),
      },
      dateBasedWorkflow: {
        resolutionDateBlank: "In Progress",
        resolutionDatePresent: "Done",
      },
    });
  } catch (error) {
    return res.status(503).json({
      success: false,
      service: "TechOps Jira API",
      projectKey: jiraProjectKey,
      issueType: jiraIssueType,
      tokenStore: {
        provider: "Upstash Redis",
        key: jiraTokenStoreKey,
        availableTokenNames: [],
        count: 0,
        fallbackConfigured: Boolean(
          String(process.env.JIRA_TOKEN || "").trim(),
        ),
        error:
          error instanceof Error ? error.message : "Token store unavailable.",
      },
      dateBasedWorkflow: {
        resolutionDateBlank: "In Progress",
        resolutionDatePresent: "Done",
      },
    });
  }
});

/*
|--------------------------------------------------------------------------
| Assignable users endpoint
|--------------------------------------------------------------------------
*/

app.get("/api/jira/assignees", async function (req, res) {
  try {
    const search = String(req.query.search || "").trim();

    const searchUrl =
      `${jiraBaseUrl}` +
      "/rest/api/2/user/assignable/search" +
      `?project=${encodeURIComponent(jiraProjectKey)}` +
      `&username=${encodeURIComponent(search)}` +
      "&maxResults=100";

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

        message: "Jira rejected the " + "assignable-user search.",

        jiraStatus: response.status,

        jiraErrors: getJiraError(jiraResult),
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
      projectKey: jiraProjectKey,
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
          : "Could not retrieve " + "Jira assignees.",
    });
  }
});

/*
|--------------------------------------------------------------------------
| Create Jira issue
|--------------------------------------------------------------------------
*/

app.post("/api/jira/issues", async function (req, res) {
  try {
    let prepared;

    try {
      prepared = await prepareIssueData(req.body);
    } catch (validationError) {
      return res.status(400).json({
        success: false,

        message:
          validationError instanceof Error
            ? validationError.message
            : "The request is invalid.",
      });
    }

    /*
     * Search Jira by Integration ID before creating.
     */
    const existingIssue = await findIssueByIntegrationId(
      prepared.cleanIntegrationId,
      prepared.jiraToken,
    );

    if (existingIssue) {
      return res.status(409).json({
        success: false,
        duplicate: true,

        jiraKey: existingIssue.key,

        jiraUrl: `${jiraBaseUrl}/browse/` + `${existingIssue.key}`,

        message:
          "This Excel row is already " + `connected to ${existingIssue.key}.`,
      });
    }

    const jiraFields = buildCreateFields(prepared);

    const jiraResponse = await fetch(`${jiraBaseUrl}/rest/api/2/issue`, {
      method: "POST",

      headers: {
        Authorization: `Bearer ${prepared.jiraToken}`,
        Accept: "application/json",
        "Content-Type": "application/json",
      },

      body: JSON.stringify({
        fields: jiraFields,
      }),
    });

    const jiraResult = await parseResponse(jiraResponse);

    if (!jiraResponse.ok) {
      console.error("Jira issue creation failed", {
        status: jiraResponse.status,
        response: jiraResult,
      });

      return res.status(jiraResponse.status).json({
        success: false,

        message: "Jira rejected the " + "issue creation request.",

        jiraStatus: jiraResponse.status,

        jiraErrors: getJiraError(jiraResult),
      });
    }

    const createdIssueKey = jiraResult.key;

    let statusResult;

    try {
      statusResult = await applyDateBasedStatus(
        createdIssueKey,
        prepared.cleanEndDate,
        prepared.jiraToken,
      );
    } catch (statusError) {
      console.error("Issue created, but status update failed", {
        issueKey: createdIssueKey,

        calculatedStatus: prepared.calculatedStatus,

        error: statusError instanceof Error ? statusError.message : statusError,
      });

      statusResult = {
        requestedStatus: prepared.calculatedStatus,

        statusChanged: false,

        actualResult: {
          success: false,
          changed: false,
        },

        message:
          statusError instanceof Error
            ? statusError.message
            : "The issue was created, " +
              "but the workflow status " +
              "could not be updated.",
      };
    }

    return res.status(201).json({
      success: true,

      jiraKey: createdIssueKey,

      jiraUrl: `${jiraBaseUrl}/browse/` + `${createdIssueKey}`,

      integrationId: prepared.cleanIntegrationId,
      jiraTokenSource: prepared.jiraTokenSource,
      jiraTokenUser: prepared.jiraTokenUser,

      assignee: prepared.assignee
        ? {
            username: prepared.assignee.name,

            displayName:
              prepared.assignee.displayName || prepared.assignee.name,
          }
        : null,

      taskDate: prepared.cleanTaskDate,

      startDate: prepared.cleanStartDate,

      endDate: prepared.cleanEndDate,

      calculatedStatus: prepared.calculatedStatus,

      statusChanged: Boolean(statusResult.statusChanged),

      statusMessage:
        statusResult.message || "No workflow status message was returned.",

      statusDetails: statusResult,

      message: statusResult.statusChanged
        ? `Jira issue ${createdIssueKey} ` +
          "was created and moved to " +
          `${prepared.calculatedStatus}.`
        : `Jira issue ${createdIssueKey} ` +
          "was created, but could not " +
          "be moved to " +
          `${prepared.calculatedStatus}.`,
    });
  } catch (error) {
    console.error("Unexpected Jira creation error", error);

    return res.status(500).json({
      success: false,

      message:
        error instanceof Error
          ? error.message
          : "The Jira integration " + "encountered an error.",
    });
  }
});

/*
|--------------------------------------------------------------------------
| Update existing Jira issue
|--------------------------------------------------------------------------
*/

app.put("/api/jira/issues/:jiraKey", async function (req, res) {
  try {
    const jiraKey = String(req.params.jiraKey || "")
      .trim()
      .toUpperCase();

    if (!validateJiraKey(jiraKey)) {
      return res.status(400).json({
        success: false,

        message: `A valid ${jiraProjectKey} ` + "Jira Key is required.",
      });
    }

    let prepared;

    try {
      prepared = await prepareIssueData(req.body);
    } catch (validationError) {
      return res.status(400).json({
        success: false,
        message:
          validationError instanceof Error
            ? validationError.message
            : "The request is invalid.",
      });
    }

    const existingIssue = await getJiraIssue(jiraKey, prepared.jiraToken);
    if (!existingIssue) {
      return res.status(404).json({
        success: false,
        message: `Jira issue ${jiraKey} was not found.`,
      });
    }

    /*
     * Ensure this Integration ID does not belong to a
     * different Jira issue.
     */
    const issueUsingIntegrationId = await findIssueByIntegrationId(
      prepared.cleanIntegrationId,
      prepared.jiraToken,
    );

    if (issueUsingIntegrationId && issueUsingIntegrationId.key !== jiraKey) {
      return res.status(409).json({
        success: false,
        duplicate: true,

        jiraKey: issueUsingIntegrationId.key,

        jiraUrl: `${jiraBaseUrl}/browse/` + `${issueUsingIntegrationId.key}`,

        message:
          "This Integration ID belongs to " +
          `${issueUsingIntegrationId.key}, ` +
          `not ${jiraKey}.`,
      });
    }

    const jiraFields = buildUpdateFields(prepared);

    const updateResponse = await fetch(
      `${jiraBaseUrl}/rest/api/2/issue/` + `${encodeURIComponent(jiraKey)}`,
      {
        method: "PUT",

        headers: {
          Authorization: `Bearer ${prepared.jiraToken}`,
          Accept: "application/json",
          "Content-Type": "application/json",
        },

        body: JSON.stringify({
          fields: jiraFields,
        }),
      },
    );

    const updateResult = await parseResponse(updateResponse);

    if (!updateResponse.ok) {
      console.error("Jira issue update failed", {
        jiraKey,
        status: updateResponse.status,
        response: updateResult,
      });

      return res.status(updateResponse.status).json({
        success: false,

        message: `Jira rejected the update ` + `for ${jiraKey}.`,

        jiraStatus: updateResponse.status,

        jiraErrors: getJiraError(updateResult),
      });
    }

    let statusResult;

    try {
      statusResult = await applyDateBasedStatus(
        jiraKey,
        prepared.cleanEndDate,
        prepared.jiraToken,
      );
    } catch (statusError) {
      console.error("Issue updated, but status change failed", {
        jiraKey,

        calculatedStatus: prepared.calculatedStatus,

        error: statusError instanceof Error ? statusError.message : statusError,
      });

      statusResult = {
        requestedStatus: prepared.calculatedStatus,

        statusChanged: false,

        actualResult: {
          success: false,
          changed: false,
        },

        message:
          statusError instanceof Error
            ? statusError.message
            : "The issue was updated, but " +
              "the workflow status could " +
              "not be changed.",
      };
    }

    return res.json({
      success: true,

      jiraKey,

      jiraUrl: `${jiraBaseUrl}/browse/${jiraKey}`,

      integrationId: prepared.cleanIntegrationId,
      jiraTokenSource: prepared.jiraTokenSource,
      jiraTokenUser: prepared.jiraTokenUser,

      assignee: prepared.assignee
        ? {
            username: prepared.assignee.name,

            displayName:
              prepared.assignee.displayName || prepared.assignee.name,
          }
        : null,

      taskDate: prepared.cleanTaskDate,

      startDate: prepared.cleanStartDate,

      endDate: prepared.cleanEndDate,

      calculatedStatus: prepared.calculatedStatus,

      statusChanged: Boolean(statusResult.statusChanged),

      statusMessage: statusResult.message || "",

      statusDetails: statusResult,

      message: statusResult.statusChanged
        ? `Jira issue ${jiraKey} was ` +
          "updated and synchronized to " +
          `${prepared.calculatedStatus}.`
        : `Jira issue ${jiraKey} was ` +
          "updated, but the workflow " +
          "status could not be moved to " +
          `${prepared.calculatedStatus}.`,
    });
  } catch (error) {
    console.error("Unexpected Jira update error", error);

    return res.status(500).json({
      success: false,

      message:
        error instanceof Error
          ? error.message
          : "The Jira update encountered " + "an internal error.",
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

  console.log(`Jira project: ${jiraProjectKey}`);

  console.log(`Jira issue type: ${jiraIssueType}`);

  console.log("Resolution Date blank: In Progress");

  console.log("Resolution Date populated: Done");
  console.log(`Jira token store: Upstash key ${jiraTokenStoreKey}`);

  console.log(
    process.env.EXPRESS_INTEGRATION_KEY
      ? "Express integration-key protection: Enabled"
      : "Express integration-key protection: Disabled",
  );
});
