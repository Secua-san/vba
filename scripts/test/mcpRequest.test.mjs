import assert from "node:assert/strict";
import test from "node:test";

import { computeBackoffDelayMs, createMcpRequestClient, McpRequestError, parseRetryAfterMs } from "../lib/mcpRequest.mjs";

function createJsonLogger() {
  const entries = [];

  return {
    entries,
    error(message) {
      entries.push(JSON.parse(message));
    },
    info(message) {
      entries.push(JSON.parse(message));
    },
    warn(message) {
      entries.push(JSON.parse(message));
    },
  };
}

test("parseRetryAfterMs supports numeric seconds and IMF-fixdate values", () => {
  const now = Date.parse("2026-03-09T00:00:00.000Z");

  assert.equal(parseRetryAfterMs("3", now), 3_000);
  assert.equal(parseRetryAfterMs("Mon, 09 Mar 2026 00:00:05 GMT", now), 5_000);
  assert.equal(parseRetryAfterMs("invalid", now), undefined);
});

test("computeBackoffDelayMs uses exponential delay and optional jitter", () => {
  assert.equal(
    computeBackoffDelayMs({
      baseDelayMs: 2_000,
      jitterRatio: 0,
      maxDelayMs: 60_000,
      random: () => 0.5,
      retryCount: 3,
    }),
    8_000,
  );

  assert.equal(
    computeBackoffDelayMs({
      baseDelayMs: 1_000,
      jitterRatio: 0.5,
      maxDelayMs: 60_000,
      random: () => 0.5,
      retryCount: 2,
    }),
    2_500,
  );
});

test("request waits for Retry-After on 429 responses and logs retry metadata", async () => {
  const sleeps = [];
  const logger = createJsonLogger();
  let attempts = 0;

  const client = createMcpRequestClient({
    fetchImpl: async () => {
      attempts += 1;
      if (attempts === 1) {
        return new Response("", {
          headers: {
            "retry-after": "3",
          },
          status: 429,
          statusText: "Too Many Requests",
        });
      }

      return new Response("ok", { status: 200 });
    },
    logger,
    maxRetries: 2,
    mcpName: "microsoft-learn",
    minIntervalMs: 0,
    sleep: async (ms) => {
      sleeps.push(ms);
    },
  });

  const result = await client.request({
    operationName: "fetch-text",
    parseResponse: (response) => response.text(),
    requestKey: "GET https://example.test/retry-after",
    url: "https://example.test/retry-after",
  });

  const retryLog = logger.entries.find((entry) => entry.event === "mcp.request.retry");

  assert.equal(result, "ok");
  assert.equal(attempts, 2);
  assert.deepEqual(sleeps, [3_000]);
  assert.deepEqual(
    retryLog,
    {
      event: "mcp.request.retry",
      finalReason: "429 Too Many Requests",
      mcpName: "microsoft-learn",
      operationName: "fetch-text",
      requestKey: "GET https://example.test/retry-after",
      retryAfterMs: 3_000,
      retryCount: 1,
      status: 429,
      url: "https://example.test/retry-after",
      waitMs: 3_000,
    },
  );
});

test("request falls back to exponential backoff when Retry-After is absent", async () => {
  const sleeps = [];
  let attempts = 0;

  const client = createMcpRequestClient({
    baseDelayMs: 1_500,
    fetchImpl: async () => {
      attempts += 1;
      if (attempts === 1) {
        return new Response("", {
          status: 429,
          statusText: "Too Many Requests",
        });
      }

      return new Response("ok", { status: 200 });
    },
    jitterRatio: 0,
    logger: createJsonLogger(),
    maxRetries: 2,
    mcpName: "microsoft-learn",
    minIntervalMs: 0,
    sleep: async (ms) => {
      sleeps.push(ms);
    },
  });

  const result = await client.request({
    operationName: "fetch-text",
    parseResponse: (response) => response.text(),
    url: "https://example.test/backoff",
  });

  assert.equal(result, "ok");
  assert.equal(attempts, 2);
  assert.deepEqual(sleeps, [1_500]);
});

test("request throttles consecutive calls by the configured interval", async () => {
  const sleeps = [];
  const startedAt = [];
  let currentTime = 0;

  const client = createMcpRequestClient({
    fetchImpl: async () => {
      startedAt.push(currentTime);
      return new Response("ok", { status: 200 });
    },
    logger: createJsonLogger(),
    mcpName: "microsoft-learn",
    minIntervalMs: 500,
    now: () => currentTime,
    sleep: async (ms) => {
      sleeps.push(ms);
      currentTime += ms;
    },
  });

  await Promise.all([
    client.request({
      operationName: "first",
      parseResponse: (response) => response.text(),
      requestKey: "first",
      url: "https://example.test/first",
    }),
    client.request({
      operationName: "second",
      parseResponse: (response) => response.text(),
      requestKey: "second",
      url: "https://example.test/second",
    }),
  ]);

  assert.deepEqual(startedAt, [0, 500]);
  assert.deepEqual(sleeps, [500]);
});

test("request deduplicates in-flight calls with the same request key", async () => {
  let fetchCount = 0;
  let releaseFetch;
  const fetchReady = new Promise((resolve) => {
    releaseFetch = resolve;
  });

  const client = createMcpRequestClient({
    fetchImpl: async () => {
      fetchCount += 1;
      await fetchReady;
      return new Response(JSON.stringify({ ok: true }), {
        headers: {
          "content-type": "application/json",
        },
        status: 200,
      });
    },
    logger: createJsonLogger(),
    mcpName: "microsoft-learn",
    minIntervalMs: 0,
  });

  const firstPromise = client.request({
    operationName: "fetch-json",
    parseResponse: (response) => response.json(),
    requestKey: "same-query",
    url: "https://example.test/dedup",
  });
  const secondPromise = client.request({
    operationName: "fetch-json",
    parseResponse: (response) => response.json(),
    requestKey: "same-query",
    url: "https://example.test/dedup",
  });

  await Promise.resolve();
  releaseFetch();

  const [firstResult, secondResult] = await Promise.all([firstPromise, secondPromise]);

  assert.equal(fetchCount, 1);
  assert.strictEqual(firstResult, secondResult);
  assert.deepEqual(firstResult, { ok: true });
});

test("request fails clearly after exceeding the maximum retry count", async () => {
  const logger = createJsonLogger();
  const sleeps = [];
  let attempts = 0;

  const client = createMcpRequestClient({
    fetchImpl: async () => {
      attempts += 1;
      return new Response("", {
        status: 429,
        statusText: "Too Many Requests",
      });
    },
    jitterRatio: 0,
    logger,
    maxRetries: 2,
    mcpName: "microsoft-learn",
    minIntervalMs: 0,
    sleep: async (ms) => {
      sleeps.push(ms);
    },
  });

  await assert.rejects(
    client.request({
      operationName: "fetch-text",
      parseResponse: (response) => response.text(),
      requestKey: "GET https://example.test/fail",
      url: "https://example.test/fail",
    }),
    (error) => {
      assert.ok(error instanceof McpRequestError);
      assert.match(error.message, /microsoft-learn/);
      assert.match(error.message, /fetch-text/);
      assert.match(error.message, /after 2 retries/);
      assert.match(error.message, /429 Too Many Requests/);
      return true;
    },
  );

  const failureLog = logger.entries.find((entry) => entry.event === "mcp.request.failed");

  assert.equal(attempts, 3);
  assert.deepEqual(sleeps, [2_000, 4_000]);
  assert.deepEqual(
    failureLog,
    {
      event: "mcp.request.failed",
      finalReason: "429 Too Many Requests",
      mcpName: "microsoft-learn",
      operationName: "fetch-text",
      requestKey: "GET https://example.test/fail",
      retryCount: 2,
      status: 429,
      url: "https://example.test/fail",
    },
  );
});
