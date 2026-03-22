const http = require("http");
const fs = require("fs");
const path = require("path");

const HOST = "127.0.0.1";
const PORT = Number(process.env.PORT || 4173);
const ROOT = __dirname;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-5-mini";

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".ico": "image/x-icon",
};

const CATEGORY_SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    rowHash: { type: "string" },
    category: { type: "string" },
    confidence: { type: "number", minimum: 0, maximum: 1 },
    reason: { type: "string" },
    shouldReview: { type: "boolean" },
    merchantKey: { type: "string" },
  },
  required: ["rowHash", "category", "confidence", "reason", "shouldReview", "merchantKey"],
};

const BATCH_SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    suggestions: {
      type: "array",
      items: CATEGORY_SCHEMA,
    },
  },
  required: ["suggestions"],
};

function sendJson(res, statusCode, body) {
  const payload = JSON.stringify(body);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(payload),
  });
  res.end(payload);
}

function sendText(res, statusCode, text) {
  res.writeHead(statusCode, {
    "Content-Type": "text/plain; charset=utf-8",
    "Content-Length": Buffer.byteLength(text),
  });
  res.end(text);
}

function safeJoin(root, requestPath) {
  const decoded = decodeURIComponent(requestPath.split("?")[0]);
  const normalized = path.normalize(decoded).replace(/^(\.\.[/\\])+/, "");
  return path.join(root, normalized);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = "";
    req.on("data", (chunk) => {
      data += chunk;
      if (data.length > 1_000_000) {
        reject(new Error("Request body too large"));
        req.destroy();
      }
    });
    req.on("end", () => resolve(data));
    req.on("error", reject);
  });
}

function extractOutputText(response) {
  if (typeof response?.output_text === "string" && response.output_text.trim()) {
    return response.output_text;
  }

  const parts = [];
  for (const item of response?.output || []) {
    for (const content of item?.content || []) {
      if (typeof content?.text === "string" && content.text.trim()) {
        parts.push(content.text);
      }
    }
  }
  return parts.join("\n").trim();
}

async function classifyTransactionsWithOpenAI(transactions, allowedCategories) {
  const promptPayload = {
    allowedCategories,
    transactions,
  };

  const apiResponse = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    body: JSON.stringify({
      model: OPENAI_MODEL,
      reasoning: { effort: "low" },
      input: [
        {
          role: "developer",
          content: [
            {
              type: "input_text",
              text:
                "You classify finance transactions into exactly one allowed category. " +
                "Be conservative. When uncertain, choose Other and set shouldReview to true. " +
                "Do not invent merchants. Use the supplied local suggestion and merchant history as hints, not facts.",
            },
          ],
        },
        {
          role: "user",
          content: [
            {
              type: "input_text",
              text: JSON.stringify(promptPayload),
            },
          ],
        },
      ],
      text: {
        format: {
          type: "json_schema",
          name: "categorization_batch",
          strict: true,
          schema: BATCH_SCHEMA,
        },
      },
    }),
  });

  if (!apiResponse.ok) {
    const errorText = await apiResponse.text();
    throw new Error(`OpenAI request failed with HTTP ${apiResponse.status}: ${errorText}`);
  }

  const responseJson = await apiResponse.json();
  const outputText = extractOutputText(responseJson);
  if (!outputText) {
    throw new Error("OpenAI response did not contain structured output text");
  }

  const parsed = JSON.parse(outputText);
  return Array.isArray(parsed?.suggestions) ? parsed.suggestions : [];
}

async function handleCategorize(req, res) {
  if (!OPENAI_API_KEY) {
    sendJson(res, 503, {
      available: false,
      message: "OPENAI_API_KEY is not configured on the server.",
    });
    return;
  }

  try {
    const rawBody = await readBody(req);
    const body = rawBody ? JSON.parse(rawBody) : {};
    const transactions = Array.isArray(body?.transactions) ? body.transactions : [];
    const allowedCategories = Array.isArray(body?.allowedCategories) ? body.allowedCategories : [];

    if (!transactions.length || !allowedCategories.length) {
      sendJson(res, 400, { error: "transactions and allowedCategories are required" });
      return;
    }

    const suggestions = await classifyTransactionsWithOpenAI(transactions, allowedCategories);
    sendJson(res, 200, {
      available: true,
      suggestions,
      model: OPENAI_MODEL,
    });
  } catch (error) {
    sendJson(res, 500, {
      available: false,
      error: error.message,
    });
  }
}

function handleAiStatus(_req, res) {
  sendJson(res, 200, {
    available: Boolean(OPENAI_API_KEY),
    model: OPENAI_MODEL,
    message: OPENAI_API_KEY
      ? `AI categorization is ready via ${OPENAI_MODEL}.`
      : "Set OPENAI_API_KEY to enable AI categorization.",
  });
}

function serveStatic(req, res) {
  const requestPath = req.url === "/" ? "/index.html" : req.url;
  const filePath = safeJoin(ROOT, requestPath);

  if (!filePath.startsWith(ROOT)) {
    sendText(res, 403, "Forbidden");
    return;
  }

  fs.stat(filePath, (statError, stats) => {
    if (statError || !stats.isFile()) {
      sendText(res, 404, "Not found");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || "application/octet-stream";
    res.writeHead(200, { "Content-Type": contentType });
    fs.createReadStream(filePath).pipe(res);
  });
}

const server = http.createServer((req, res) => {
  if (!req.url) {
    sendText(res, 400, "Missing URL");
    return;
  }

  if (req.method === "GET" && req.url.startsWith("/api/ai-status")) {
    handleAiStatus(req, res);
    return;
  }

  if (req.method === "POST" && req.url.startsWith("/api/categorize")) {
    handleCategorize(req, res);
    return;
  }

  if (req.method === "GET" || req.method === "HEAD") {
    serveStatic(req, res);
    return;
  }

  sendText(res, 405, "Method not allowed");
});

server.listen(PORT, HOST, () => {
  console.log(`Finance tool server running at http://${HOST}:${PORT}`);
});
