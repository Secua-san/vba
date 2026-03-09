const DEFAULT_BASE_DELAY_MS = 2_000;
const DEFAULT_JITTER_RATIO = 0.2;
const DEFAULT_MAX_DELAY_MS = 60_000;
const DEFAULT_MAX_RETRIES = 5;
const DEFAULT_MIN_INTERVAL_MS = 250;
const DEFAULT_TIMEOUT_MS = 30_000;

export class McpRequestError extends Error {
  constructor(message, { cause, finalReason, mcpName, operationName, retryCount } = {}) {
    super(message, cause ? { cause } : undefined);
    this.name = "McpRequestError";
    this.finalReason = finalReason;
    this.mcpName = mcpName;
    this.operationName = operationName;
    this.retryCount = retryCount;
  }
}

export function parseRetryAfterMs(retryAfterHeader, now = Date.now()) {
  if (!retryAfterHeader) {
    return undefined;
  }

  const seconds = Number(retryAfterHeader);
  if (Number.isFinite(seconds) && seconds >= 0) {
    return Math.round(seconds * 1_000);
  }

  const retryDate = Date.parse(retryAfterHeader);
  if (Number.isFinite(retryDate)) {
    const delayMs = retryDate - now;
    return delayMs > 0 ? delayMs : undefined;
  }

  return undefined;
}

export function computeBackoffDelayMs({
  baseDelayMs = DEFAULT_BASE_DELAY_MS,
  jitterRatio = DEFAULT_JITTER_RATIO,
  maxDelayMs = DEFAULT_MAX_DELAY_MS,
  random = Math.random,
  retryCount,
}) {
  const normalizedRetryCount = Math.max(1, retryCount);
  const baseDelay = Math.min(maxDelayMs, baseDelayMs * 2 ** (normalizedRetryCount - 1));

  if (jitterRatio <= 0) {
    return baseDelay;
  }

  const jitterSpan = Math.max(0, Math.round(baseDelay * jitterRatio));
  const jitter = Math.floor(random() * (jitterSpan + 1));
  return Math.min(maxDelayMs, baseDelay + jitter);
}

function defaultSleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function defaultShouldRetryResponse(response) {
  return response.status === 429;
}

function defaultShouldRetryError(error) {
  return error?.name === "AbortError" || error?.name === "TimeoutError" || error instanceof TypeError;
}

function describeResponseReason(response) {
  return `${response.status} ${response.statusText || "Unknown Error"}`.trim();
}

function describeError(error) {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

function createLogger(logger, mcpName) {
  return (level, event, payload = {}) => {
    const sink =
      typeof logger?.[level] === "function"
        ? logger[level].bind(logger)
        : typeof console[level] === "function"
          ? console[level].bind(console)
          : console.log.bind(console);

    sink(
      JSON.stringify({
        event,
        mcpName,
        ...payload,
      }),
    );
  };
}

function buildRequestKey(url, init, requestKey) {
  if (requestKey) {
    return requestKey;
  }

  const method = String(init.method || "GET").toUpperCase();
  const body = typeof init.body === "string" && init.body.length > 0 ? ` ${init.body}` : "";
  return `${method} ${url}${body}`;
}

function withTimeoutSignal(init, timeoutMs) {
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0 || init.signal) {
    return init;
  }

  return {
    ...init,
    signal: AbortSignal.timeout(timeoutMs),
  };
}

function buildFailureError({ cause, finalReason, mcpName, operationName, retryCount }) {
  const retrySuffix = retryCount === 0 ? "without retries" : `after ${retryCount} retries`;
  return new McpRequestError(`MCP request to ${mcpName} failed for ${operationName} ${retrySuffix}: ${finalReason}`, {
    cause,
    finalReason,
    mcpName,
    operationName,
    retryCount,
  });
}

export function createMcpRequestClient({
  baseDelayMs = DEFAULT_BASE_DELAY_MS,
  fetchImpl = globalThis.fetch,
  jitterRatio = DEFAULT_JITTER_RATIO,
  logger = console,
  maxDelayMs = DEFAULT_MAX_DELAY_MS,
  maxRetries = DEFAULT_MAX_RETRIES,
  mcpName,
  minIntervalMs = DEFAULT_MIN_INTERVAL_MS,
  now = Date.now,
  random = Math.random,
  shouldRetryError = defaultShouldRetryError,
  shouldRetryResponse = defaultShouldRetryResponse,
  sleep = defaultSleep,
  timeoutMs = DEFAULT_TIMEOUT_MS,
} = {}) {
  if (!mcpName) {
    throw new Error("mcpName is required.");
  }

  if (typeof fetchImpl !== "function") {
    throw new Error("fetchImpl must be a function.");
  }

  const log = createLogger(logger, mcpName);
  const inFlightRequests = new Map();
  let nextAvailableAt = 0;
  let rateLimitQueue = Promise.resolve();

  function waitForRateLimitSlot({ operationName, requestKey, url }) {
    const queued = rateLimitQueue.then(async () => {
      const waitMs = Math.max(0, nextAvailableAt - now());
      if (waitMs > 0) {
        log("info", "mcp.rate_limit.wait", {
          operationName,
          requestKey,
          url,
          waitMs,
        });
        await sleep(waitMs);
      }

      nextAvailableAt = now() + minIntervalMs;
    });

    rateLimitQueue = queued.catch(() => undefined);
    return queued;
  }

  async function executeRequest({ init = {}, operationName = "request", parseResponse, requestKey, url }) {
    let finalReason = "Unknown failure";

    for (let attempt = 0; attempt <= maxRetries; attempt += 1) {
      await waitForRateLimitSlot({ operationName, requestKey, url });

      let response;
      try {
        response = await fetchImpl(url, withTimeoutSignal(init, timeoutMs));
      } catch (error) {
        finalReason = describeError(error);

        if (attempt >= maxRetries || !shouldRetryError(error)) {
          log("error", "mcp.request.failed", {
            finalReason,
            operationName,
            requestKey,
            retryCount: attempt,
            url,
          });
          throw buildFailureError({
            cause: error,
            finalReason,
            mcpName,
            operationName,
            retryCount: attempt,
          });
        }

        const waitMs = computeBackoffDelayMs({
          baseDelayMs,
          jitterRatio,
          maxDelayMs,
          random,
          retryCount: attempt + 1,
        });

        log("warn", "mcp.request.retry", {
          finalReason,
          operationName,
          requestKey,
          retryCount: attempt + 1,
          url,
          waitMs,
        });
        await sleep(waitMs);
        continue;
      }

      if (response.ok) {
        try {
          return await parseResponse(response);
        } catch (error) {
          finalReason = `Failed to parse response: ${describeError(error)}`;
          log("error", "mcp.request.failed", {
            finalReason,
            operationName,
            requestKey,
            retryCount: attempt,
            url,
          });
          throw buildFailureError({
            cause: error,
            finalReason,
            mcpName,
            operationName,
            retryCount: attempt,
          });
        }
      }

      finalReason = describeResponseReason(response);

      if (attempt >= maxRetries || !shouldRetryResponse(response)) {
        log("error", "mcp.request.failed", {
          finalReason,
          operationName,
          requestKey,
          retryCount: attempt,
          status: response.status,
          url,
        });
        throw buildFailureError({
          finalReason,
          mcpName,
          operationName,
          retryCount: attempt,
        });
      }

      const retryAfterMs = parseRetryAfterMs(response.headers.get("retry-after"), now());
      const waitMs =
        retryAfterMs ??
        computeBackoffDelayMs({
          baseDelayMs,
          jitterRatio,
          maxDelayMs,
          random,
          retryCount: attempt + 1,
        });

      log("warn", "mcp.request.retry", {
        finalReason,
        operationName,
        requestKey,
        retryCount: attempt + 1,
        retryAfterMs,
        status: response.status,
        url,
        waitMs,
      });
      await sleep(waitMs);
    }

    log("error", "mcp.request.failed", {
      finalReason,
      operationName,
      requestKey,
      retryCount: maxRetries,
      url,
    });
    throw buildFailureError({
      finalReason,
      mcpName,
      operationName,
      retryCount: maxRetries,
    });
  }

  async function request({ init = {}, operationName = "request", parseResponse, requestKey, url }) {
    if (!url) {
      throw new Error("url is required.");
    }

    if (typeof parseResponse !== "function") {
      throw new Error("parseResponse must be a function.");
    }

    const normalizedRequestKey = buildRequestKey(url, init, requestKey);
    const existing = inFlightRequests.get(normalizedRequestKey);

    if (existing) {
      log("info", "mcp.request.deduplicated", {
        operationName,
        requestKey: normalizedRequestKey,
        url,
      });
      return existing;
    }

    const promise = executeRequest({
      init,
      operationName,
      parseResponse,
      requestKey: normalizedRequestKey,
      url,
    }).finally(() => {
      inFlightRequests.delete(normalizedRequestKey);
    });

    inFlightRequests.set(normalizedRequestKey, promise);
    return promise;
  }

  return {
    request,
  };
}
