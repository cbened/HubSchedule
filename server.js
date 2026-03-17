const path = require("path");
const express = require("express");
const dotenv = require("dotenv");

dotenv.config();

const app = express();
const port = Number(process.env.PORT || 3000);
const formId = process.env.HUBSPOT_FORM_ID;
const portalId = process.env.HUBSPOT_PORTAL_ID;
const token = process.env.HUBSPOT_PRIVATE_APP_TOKEN;
const pageSize = Number(process.env.SUBMISSION_PAGE_SIZE || 25);

app.use(express.static(path.join(__dirname, "public")));

app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

app.get("/api/submissions", async (_req, res) => {
  if (!formId || !token) {
    return res.status(500).json({
      error: "Missing HUBSPOT_FORM_ID or HUBSPOT_PRIVATE_APP_TOKEN in environment.",
    });
  }

  try {
    const endpoint = `https://api.hubapi.com/form-integrations/v1/submissions/forms/${formId}`;
    const query = new URLSearchParams({ limit: String(pageSize) });
    const response = await fetch(`${endpoint}?${query.toString()}`, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (!response.ok) {
      const body = await response.text();
      return res.status(response.status).json({
        error: `HubSpot API error: ${body}`,
      });
    }

    const data = await response.json();
    const results = normalizeSubmissions(data?.results || data || []);

    return res.json({
      meta: {
        formId,
        portalId: portalId || null,
      },
      results,
    });
  } catch (error) {
    return res.status(500).json({
      error: error.message || "Unexpected server error.",
    });
  }
});

function normalizeSubmissions(records) {
  return records.map((row) => {
    const values = {};
    const fields = row.values || row.submissionValues || [];

    if (Array.isArray(fields)) {
      for (const field of fields) {
        const key = field.name || field.fieldName || "field";
        values[key] = field.value;
      }
    } else {
      Object.assign(values, fields);
    }

    return {
      submittedAt: row.submittedAt || row.submittedAtTimestamp || row.timestamp || Date.now(),
      values,
    };
  });
}

app.listen(port, () => {
  console.log(`HubSchedule add-in server listening on http://localhost:${port}`);
});
