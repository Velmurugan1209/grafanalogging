const HTTP_METHODS = "GET|POST|PUT|PATCH|DELETE";

export function cleanLogText(log: string): string {
  return log.replace(/\x1b\[[0-9;]*m/g, "");
}

export function resolveDetectLevel(
  detectedLevel: string,
  cleanLog: string
): string {
  if (detectedLevel && detectedLevel !== "unknown") {
    return detectedLevel;
  }

  const levelMatch = cleanLog.match(/\b(error|warn|fatal|debug):\s/i);
  return levelMatch?.[1]?.toLowerCase() ?? detectedLevel;
}

function parseJsonMetadata(cleanLog: string): Record<string, unknown> | null {
  const jsonStart = cleanLog.indexOf("{");
  if (jsonStart < 0) return null;

  try {
    const parsed = JSON.parse(cleanLog.slice(jsonStart));
    return typeof parsed === "object" && parsed !== null
      ? (parsed as Record<string, unknown>)
      : null;
  } catch {
    return null;
  }
}

function normalizeEndpoint(endpoint: string): string {
  return endpoint.replace(/\/\d+$/, "");
}

export function extractMethodAndEndpoint(cleanLog: string): {
  method: string;
  endpoint: string;
} {
  const meta = parseJsonMetadata(cleanLog);

  let method = "";
  let endpoint = "";

  if (meta) {
    if (typeof meta.method === "string") {
      method = meta.method;
    }

    if (typeof meta.path === "string") {
      endpoint = meta.path;
    }

    if (!endpoint && typeof meta.endpoint === "string") {
      const endpointValue = meta.endpoint;
      const endpointParts = endpointValue.match(
        new RegExp(`^(${HTTP_METHODS})\\s+(\\S+)`)
      );
      if (endpointParts) {
        if (!method) method = endpointParts[1];
        endpoint = endpointParts[2];
      }
    }
  }

  if (!method) {
    const methodMatch =
      cleanLog.match(new RegExp(`"(${HTTP_METHODS})\\s+[^"]+"`)) ??
      cleanLog.match(new RegExp(`"endpoint":"(${HTTP_METHODS})\\s+`)) ??
      cleanLog.match(new RegExp(`"method":"(${HTTP_METHODS})"`));

    if (methodMatch) {
      method = methodMatch[1] ?? "";
    } else {
      const bracketMatches = [
        ...cleanLog.matchAll(new RegExp(`\\[(${HTTP_METHODS})\\s+\\/`, "g")),
      ];
      method = bracketMatches[bracketMatches.length - 1]?.[1] ?? "";
    }
  }

  if (!endpoint) {
    const endpointMatch =
      cleanLog.match(new RegExp(`"(${HTTP_METHODS})\\s+(\\/[^\\s?"]+)`)) ??
      cleanLog.match(
        new RegExp(`"endpoint":"(?:${HTTP_METHODS})\\s+(\\/[^\\s?"]+)`)
      ) ??
      cleanLog.match(/"path":"(\/[^"?]+)/);

    if (endpointMatch) {
      endpoint = endpointMatch[2] ?? endpointMatch[1] ?? "";
    } else {
      const bracketMatches = [
        ...cleanLog.matchAll(
          new RegExp(`\\[(?:${HTTP_METHODS})\\s+(\\/[^\\]]+)\\]`, "g")
        ),
      ];
      const lastBracket = bracketMatches[bracketMatches.length - 1];
      endpoint = lastBracket?.[1] ?? "";
    }
  }

  return {
    method,
    endpoint: normalizeEndpoint(endpoint),
  };
}

export function extractErrorMessage(cleanLog: string): string {
  const errorMsgMatch = cleanLog.match(
    /(?:error|warn|fatal|debug):\s*([^{]*)/i
  );
  return errorMsgMatch?.[1]?.trim() ?? "";
}

export function extractDate(cleanLog: string): string {
  const tsJson = cleanLog.match(/"timestamp":"([^"]+)"/);
  let date = tsJson?.[1] ?? "";

  if (!date) {
    const tsIso = cleanLog.match(/^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})/);
    if (tsIso) date = tsIso[1];
  }

  const isoPrefix = date.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoPrefix) {
    return isoPrefix[1];
  }

  const apacheDate = date.match(/^(\d{2})\/(\w{3})\/(\d{4})/);
  if (apacheDate) {
    const months: Record<string, string> = {
      Jan: "01",
      Feb: "02",
      Mar: "03",
      Apr: "04",
      May: "05",
      Jun: "06",
      Jul: "07",
      Aug: "08",
      Sep: "09",
      Oct: "10",
      Nov: "11",
      Dec: "12",
    };
    return `${apacheDate[3]}-${months[apacheDate[2]] ?? "01"}-${apacheDate[1]}`;
  }

  return date;
}
