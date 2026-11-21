import test from "node:test";
import assert from "node:assert/strict";

import {
  explainCorrectionServiceDecision,
  explainLemmatizerServiceDecision,
  resolveCorrectionServiceDecision,
  resolveLemmatizerServiceDecision,
} from "../src/config/serviceConfig.js";

test("correction service defaults to safe mock mode", () => {
  const decision = resolveCorrectionServiceDecision({ env: {}, windowObject: undefined });

  assert.equal(decision.provider, "mock");
  assert.equal(decision.usesRemote, false);
  assert.equal(decision.isAllowed, true);
  assert.equal(decision.blockReason, null);
});

test("legacy mock flag still maps to remote when explicitly disabled", () => {
  const decision = resolveCorrectionServiceDecision({
    env: { VEJICE_USE_MOCK: "false" },
    windowObject: undefined,
  });

  assert.equal(decision.provider, "remote");
  assert.equal(decision.usesRemote, true);
  assert.equal(decision.isAllowed, false);
  assert.equal(decision.blockReason, "unsafe_remote_correction_api_blocked");
});

test("trusted remote correction service is allowed", () => {
  const decision = resolveCorrectionServiceDecision({
    env: {
      VEJICE_CORRECTION_PROVIDER: "remote",
      VEJICE_REMOTE_CORRECTION_TRUSTED: "true",
    },
    windowObject: undefined,
  });

  assert.equal(decision.provider, "remote");
  assert.equal(decision.remoteTrusted, true);
  assert.equal(decision.isAllowed, true);
});

test("unsafe legacy correction service requires explicit opt-in", () => {
  const decision = resolveCorrectionServiceDecision({
    env: {
      VEJICE_CORRECTION_PROVIDER: "remote",
      VEJICE_ALLOW_UNSAFE_REMOTE_API: "true",
    },
    windowObject: undefined,
  });

  assert.equal(decision.provider, "remote");
  assert.equal(decision.allowUnsafe, true);
  assert.equal(decision.isAllowed, true);
});

test("lemmatizer defaults to synthetic mode", () => {
  const decision = resolveLemmatizerServiceDecision({ env: {}, windowObject: undefined });

  assert.equal(decision.provider, "synthetic");
  assert.equal(decision.usesRemote, false);
  assert.equal(decision.isAllowed, true);
});

test("legacy lemmatizer flag remains supported but blocked unless approved", () => {
  const decision = resolveLemmatizerServiceDecision({
    env: { VEJICE_USE_LEMMATIZER: "true" },
    windowObject: undefined,
  });

  assert.equal(decision.provider, "remote");
  assert.equal(decision.isAllowed, false);
  assert.match(explainLemmatizerServiceDecision(decision), /blocked by default/i);
});

test("blocked correction decision explains the trust gate", () => {
  const decision = resolveCorrectionServiceDecision({
    env: { VEJICE_CORRECTION_PROVIDER: "remote" },
    windowObject: undefined,
  });

  assert.match(explainCorrectionServiceDecision(decision), /blocked by default/i);
});
