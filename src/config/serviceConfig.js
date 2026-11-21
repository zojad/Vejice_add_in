/* global process, window */

const BOOLEAN_TRUE = new Set(["1", "true", "yes", "on"]);
const BOOLEAN_FALSE = new Set(["0", "false", "no", "off"]);
const CORRECTION_PROVIDERS = new Set(["mock", "remote"]);
const LEMMATIZER_PROVIDERS = new Set(["synthetic", "remote"]);

function getBuildEnv() {
  return {
    VEJICE_CORRECTION_PROVIDER:
      typeof __VEJICE_BUILD_CORRECTION_PROVIDER__ !== "undefined"
        ? __VEJICE_BUILD_CORRECTION_PROVIDER__
        : undefined,
    VEJICE_REMOTE_CORRECTION_TRUSTED:
      typeof __VEJICE_BUILD_REMOTE_CORRECTION_TRUSTED__ !== "undefined"
        ? __VEJICE_BUILD_REMOTE_CORRECTION_TRUSTED__
        : undefined,
    VEJICE_ALLOW_UNSAFE_REMOTE_API:
      typeof __VEJICE_BUILD_ALLOW_UNSAFE_REMOTE_API__ !== "undefined"
        ? __VEJICE_BUILD_ALLOW_UNSAFE_REMOTE_API__
        : undefined,
    VEJICE_USE_MOCK:
      typeof __VEJICE_BUILD_USE_MOCK__ !== "undefined" ? __VEJICE_BUILD_USE_MOCK__ : undefined,
    VEJICE_LEMMATIZER_PROVIDER:
      typeof __VEJICE_BUILD_LEMMATIZER_PROVIDER__ !== "undefined"
        ? __VEJICE_BUILD_LEMMATIZER_PROVIDER__
        : undefined,
    VEJICE_REMOTE_LEMMATIZER_TRUSTED:
      typeof __VEJICE_BUILD_REMOTE_LEMMATIZER_TRUSTED__ !== "undefined"
        ? __VEJICE_BUILD_REMOTE_LEMMATIZER_TRUSTED__
        : undefined,
    VEJICE_ALLOW_UNSAFE_LEMMATIZER:
      typeof __VEJICE_BUILD_ALLOW_UNSAFE_LEMMATIZER__ !== "undefined"
        ? __VEJICE_BUILD_ALLOW_UNSAFE_LEMMATIZER__
        : undefined,
    VEJICE_USE_LEMMATIZER:
      typeof __VEJICE_BUILD_USE_LEMMATIZER__ !== "undefined"
        ? __VEJICE_BUILD_USE_LEMMATIZER__
        : undefined,
    VEJICE_LEMMATIZER_ANCHORS:
      typeof __VEJICE_BUILD_LEMMATIZER_ANCHORS__ !== "undefined"
        ? __VEJICE_BUILD_LEMMATIZER_ANCHORS__
        : undefined,
  };
}

function getRuntime(runtime = {}) {
  const buildEnv = getBuildEnv();
  const runtimeEnv =
    runtime && typeof runtime === "object" && runtime.env
      ? runtime.env
      : typeof process !== "undefined" && process?.env
        ? process.env
        : {};
  return {
    // Explicit runtime values remain highest priority (for tests and controlled overrides).
    env: { ...buildEnv, ...runtimeEnv },
    windowObject:
      runtime && typeof runtime === "object" && "windowObject" in runtime
        ? runtime.windowObject
        : typeof window !== "undefined"
          ? window
          : undefined,
  };
}

export function parseBooleanFlag(value) {
  if (typeof value === "boolean") return value;
  if (typeof value !== "string") return undefined;
  const normalized = value.trim().toLowerCase();
  if (!normalized) return undefined;
  if (BOOLEAN_TRUE.has(normalized)) return true;
  if (BOOLEAN_FALSE.has(normalized)) return false;
  return undefined;
}

function readWindowString(windowObject, key) {
  if (!windowObject || typeof key !== "string") return undefined;
  const value = windowObject[key];
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed || undefined;
}

function readEnvString(env, key) {
  if (!env || typeof key !== "string") return undefined;
  const value = env[key];
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed || undefined;
}

function readBooleanFlag(windowObject, env, { windowKeys = [], envKeys = [] } = {}) {
  for (const key of windowKeys) {
    if (!key) continue;
    const parsed = parseBooleanFlag(windowObject?.[key]);
    if (typeof parsed === "boolean") return parsed;
  }
  for (const key of envKeys) {
    if (!key) continue;
    const parsed = parseBooleanFlag(env?.[key]);
    if (typeof parsed === "boolean") return parsed;
  }
  return undefined;
}

function normalizeCorrectionProvider(value) {
  if (typeof value !== "string") return undefined;
  const normalized = value.trim().toLowerCase();
  if (normalized === "legacy-remote") return "remote";
  return CORRECTION_PROVIDERS.has(normalized) ? normalized : undefined;
}

function normalizeLemmatizerProvider(value) {
  if (typeof value !== "string") return undefined;
  const normalized = value.trim().toLowerCase();
  return LEMMATIZER_PROVIDERS.has(normalized) ? normalized : undefined;
}

export function resolveCorrectionServiceDecision(runtime = {}) {
  const { env, windowObject } = getRuntime(runtime);
  const explicitProvider =
    normalizeCorrectionProvider(readWindowString(windowObject, "__VEJICE_CORRECTION_PROVIDER__")) ||
    normalizeCorrectionProvider(readEnvString(env, "VEJICE_CORRECTION_PROVIDER"));
  const legacyMockFlag = readBooleanFlag(windowObject, env, {
    windowKeys: ["__VEJICE_USE_MOCK__"],
    envKeys: ["VEJICE_USE_MOCK"],
  });
  const provider =
    explicitProvider ||
    (typeof legacyMockFlag === "boolean" ? (legacyMockFlag ? "mock" : "remote") : "mock");
  const remoteTrusted =
    readBooleanFlag(windowObject, env, {
      windowKeys: ["__VEJICE_REMOTE_CORRECTION_TRUSTED__"],
      envKeys: ["VEJICE_REMOTE_CORRECTION_TRUSTED"],
    }) ?? false;
  const allowUnsafe =
    readBooleanFlag(windowObject, env, {
      windowKeys: ["__VEJICE_ALLOW_UNSAFE_REMOTE_API__"],
      envKeys: ["VEJICE_ALLOW_UNSAFE_REMOTE_API"],
    }) ?? false;
  const usesRemote = provider === "remote";
  const isAllowed = !usesRemote || remoteTrusted || allowUnsafe;
  return {
    provider,
    usesRemote,
    remoteTrusted,
    allowUnsafe,
    isAllowed,
    blockReason: usesRemote && !isAllowed ? "unsafe_remote_correction_api_blocked" : null,
  };
}

export function resolveLemmatizerServiceDecision(runtime = {}) {
  const { env, windowObject } = getRuntime(runtime);
  const explicitProvider =
    normalizeLemmatizerProvider(
      readWindowString(windowObject, "__VEJICE_LEMMATIZER_PROVIDER__")
    ) ||
    normalizeLemmatizerProvider(readEnvString(env, "VEJICE_LEMMATIZER_PROVIDER"));
  const legacyUseFlag = readBooleanFlag(windowObject, env, {
    windowKeys: ["__VEJICE_USE_LEMMATIZER__"],
    envKeys: ["VEJICE_USE_LEMMATIZER", "VEJICE_LEMMATIZER_ANCHORS"],
  });
  const provider =
    explicitProvider ||
    (typeof legacyUseFlag === "boolean" ? (legacyUseFlag ? "remote" : "synthetic") : "synthetic");
  const remoteTrusted =
    readBooleanFlag(windowObject, env, {
      windowKeys: ["__VEJICE_REMOTE_LEMMATIZER_TRUSTED__"],
      envKeys: ["VEJICE_REMOTE_LEMMATIZER_TRUSTED"],
    }) ?? false;
  const allowUnsafe =
    readBooleanFlag(windowObject, env, {
      windowKeys: ["__VEJICE_ALLOW_UNSAFE_LEMMATIZER__"],
      envKeys: ["VEJICE_ALLOW_UNSAFE_LEMMATIZER"],
    }) ?? false;
  const usesRemote = provider === "remote";
  const isAllowed = !usesRemote || remoteTrusted || allowUnsafe;
  return {
    provider,
    usesRemote,
    remoteTrusted,
    allowUnsafe,
    isAllowed,
    blockReason: usesRemote && !isAllowed ? "unsafe_remote_lemmatizer_blocked" : null,
  };
}

export function explainCorrectionServiceDecision(decision = resolveCorrectionServiceDecision()) {
  if (!decision?.blockReason) return "";
  return "Remote comma-correction API is blocked by default. Set VEJICE_REMOTE_CORRECTION_TRUSTED=true for a trusted service or VEJICE_ALLOW_UNSAFE_REMOTE_API=true to opt into the legacy unsafe API.";
}

export function explainLemmatizerServiceDecision(decision = resolveLemmatizerServiceDecision()) {
  if (!decision?.blockReason) return "";
  return "Remote lemmatizer is blocked by default. Set VEJICE_REMOTE_LEMMATIZER_TRUSTED=true for a trusted service or VEJICE_ALLOW_UNSAFE_LEMMATIZER=true to opt into the legacy unsafe lemmatizer.";
}

export function describeInternalTestingMode(runtime = {}) {
  const correction = resolveCorrectionServiceDecision(runtime);
  const lemmatizer = resolveLemmatizerServiceDecision(runtime);
  const correctionLabel =
    correction.provider === "mock"
      ? "mock popravljanje"
      : correction.remoteTrusted
        ? "zaupan oddaljeni popravljalnik"
        : correction.allowUnsafe
          ? "nezaupan oddaljeni popravljalnik"
          : "blokiran oddaljeni popravljalnik";
  const lemmatizerLabel =
    lemmatizer.provider === "synthetic"
      ? "sintetično sidranje"
      : lemmatizer.remoteTrusted
        ? "zaupan oddaljeni lematizator"
        : lemmatizer.allowUnsafe
          ? "nezaupan oddaljeni lematizator"
          : "blokiran oddaljeni lematizator";
  return `${correctionLabel}; ${lemmatizerLabel}`;
}
