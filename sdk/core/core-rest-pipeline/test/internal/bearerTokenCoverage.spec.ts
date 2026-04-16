// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { AccessToken, TokenCredential } from "@azure/core-auth";
import type {
  AuthorizeRequestOnChallengeOptions,
  PipelineResponse,
  SendRequest,
} from "../../src/index.js";
import {
  bearerTokenAuthenticationPolicy,
  createHttpHeaders,
  createPipelineRequest,
} from "../../src/index.js";
import { describe, it, assert, expect, vi, beforeEach, afterEach } from "vitest";

describe("BearerTokenAuthenticationPolicy - additional coverage", function () {
  beforeEach(() => {
    vi.useFakeTimers({ now: Date.now() });
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  it("returns response when second CAE challenge after custom handler has unparsable claims", async function () {
    const tokenExpiration = Date.now() + 1000 * 60;
    const getToken = vi.fn<() => Promise<AccessToken | null>>();
    getToken.mockResolvedValue({
      token: "token",
      expiresOnTimestamp: tokenExpiration,
    });
    const credential: TokenCredential = { getToken };
    const scopes = ["test-scope"];
    const request = createPipelineRequest({ url: "https://example.com" });

    // First response: non-CAE challenge (handled by custom callback)
    const nonCaeChallengeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer authorization_uri="https://login.windows.net/", error="invalid_token"`,
      }),
      request,
      status: 401,
    };

    // Second response after custom handler: CAE challenge with unparsable base64 claims
    const caeChallengeWithBadClaims: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer error="insufficient_claims", claims="not valid base64!!!"`,
      }),
      request,
      status: 401,
    };

    const next = vi.fn<SendRequest>();
    next
      .mockResolvedValueOnce(nonCaeChallengeResponse) // initial request -> non-CAE challenge
      .mockResolvedValueOnce(caeChallengeWithBadClaims); // after custom handler -> CAE with bad claims

    async function authorizeRequestOnChallenge(
      options: AuthorizeRequestOnChallengeOptions,
    ): Promise<boolean> {
      const token = await options.getAccessToken(scopes, {});
      if (token) {
        options.request.headers.set("Authorization", `Bearer ${token.token}`);
        return true;
      }
      return false;
    }

    const policy = bearerTokenAuthenticationPolicy({
      scopes,
      credential,
      challengeCallbacks: { authorizeRequestOnChallenge },
    });

    // Should return the response (lines 291-294: unparsable claims after custom handler)
    const response = await policy.sendRequest(request, next);
    assert.equal(response.status, 401);
  });

  it("handles second CAE challenge after custom handler with valid claims", async function () {
    const tokenExpiration = Date.now() + 1000 * 60;
    const getToken = vi.fn<() => Promise<AccessToken | null>>();
    // Initial token
    getToken.mockResolvedValueOnce({
      token: "initial-token",
      expiresOnTimestamp: tokenExpiration,
    });
    // Token for custom handler
    getToken.mockResolvedValueOnce({
      token: "custom-handler-token",
      expiresOnTimestamp: tokenExpiration,
    });
    // Token for CAE challenge
    getToken.mockResolvedValueOnce({
      token: "cae-token",
      expiresOnTimestamp: tokenExpiration,
    });

    const credential: TokenCredential = { getToken };
    const scopes = ["test-scope"];
    const request = createPipelineRequest({ url: "https://example.com" });

    // First: non-CAE challenge
    const nonCaeChallengeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer authorization_uri="https://login.windows.net/", error="invalid_token"`,
      }),
      request,
      status: 401,
    };

    // After custom handler: valid CAE challenge (base64 of '{"test":"value"}' = eyJ0ZXN0IjoidmFsdWUifQ==)
    const validCaeChallengeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer error="insufficient_claims", claims="eyJ0ZXN0IjoidmFsdWUifQ=="`,
      }),
      request,
      status: 401,
    };

    const successResponse: PipelineResponse = {
      headers: createHttpHeaders(),
      request,
      status: 200,
    };

    const next = vi.fn<SendRequest>();
    next
      .mockResolvedValueOnce(nonCaeChallengeResponse)
      .mockResolvedValueOnce(validCaeChallengeResponse)
      .mockResolvedValueOnce(successResponse);

    async function authorizeRequestOnChallenge(
      options: AuthorizeRequestOnChallengeOptions,
    ): Promise<boolean> {
      const token = await options.getAccessToken(scopes, {});
      if (token) {
        options.request.headers.set("Authorization", `Bearer ${token.token}`);
        return true;
      }
      return false;
    }

    const policy = bearerTokenAuthenticationPolicy({
      scopes,
      credential,
      challengeCallbacks: { authorizeRequestOnChallenge },
    });

    const response = await policy.sendRequest(request, next);
    assert.equal(response.status, 200);
    // 3 calls: initial + custom handler + CAE handler
    expect(next).toHaveBeenCalledTimes(3);
  });

  it("handles second CAE challenge after custom handler when no credential provided", async function () {
    const scopes = ["test-scope"];
    const request = createPipelineRequest({ url: "https://example.com" });

    const nonCaeChallengeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer authorization_uri="https://login.windows.net/", error="invalid_token"`,
      }),
      request,
      status: 401,
    };

    // Valid CAE challenge
    const validCaeChallengeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer error="insufficient_claims", claims="eyJ0ZXN0IjoidmFsdWUifQ=="`,
      }),
      request,
      status: 401,
    };

    const next = vi.fn<SendRequest>();
    next
      .mockResolvedValueOnce(nonCaeChallengeResponse)
      .mockResolvedValueOnce(validCaeChallengeResponse);

    async function authorizeRequestOnChallenge(
      _options: AuthorizeRequestOnChallengeOptions,
    ): Promise<boolean> {
      // Just return true to re-send the request
      return true;
    }

    // No credential → getAccessToken returns null → authorizeRequestOnCaeChallenge returns false
    const policy = bearerTokenAuthenticationPolicy({
      scopes,
      credential: undefined,
      challengeCallbacks: { authorizeRequestOnChallenge },
    });

    const response = await policy.sendRequest(request, next);
    assert.equal(response.status, 401);
    expect(next).toHaveBeenCalledTimes(2);
  });

  it("skips CAE handling when WWW-Authenticate header value is empty", async function () {
    // This test covers getCaeChallengeClaims line 377: early return when challenges is falsy.
    // An empty header value passes isChallengeResponse (has() returns true) but
    // getCaeChallengeClaims("") returns undefined since "" is falsy.
    const tokenExpiration = Date.now() + 1000 * 60;
    const credential: TokenCredential = {
      getToken: async () => ({
        token: "token",
        expiresOnTimestamp: tokenExpiration,
      }),
    };
    const scopes = ["test-scope"];
    const request = createPipelineRequest({ url: "https://example.com" });

    // 401 with empty WWW-Authenticate
    const challengeResponse: PipelineResponse = {
      headers: createHttpHeaders({ "WWW-Authenticate": "" }),
      request,
      status: 401,
    };

    const next = vi.fn<SendRequest>();
    next.mockResolvedValue(challengeResponse);

    const policy = bearerTokenAuthenticationPolicy({ scopes, credential });

    const response = await policy.sendRequest(request, next);
    // Should just return the 401 response since getCaeChallengeClaims returns undefined
    assert.equal(response.status, 401);
    // Only 1 call - no retry since claims is undefined
    expect(next).toHaveBeenCalledOnce();
  });

  it("does not retry when authorizeRequestOnCaeChallenge returns false", async function () {
    // Covers line 265: shouldSendRequest = false after first CAE challenge
    const scopes = ["test-scope"];
    const request = createPipelineRequest({ url: "https://example.com" });

    // CAE challenge response with valid base64 claims
    const caeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer error="insufficient_claims", claims="eyJ0ZXN0IjoidmFsdWUifQ=="`,
      }),
      request,
      status: 401,
    };

    const next = vi.fn<SendRequest>();
    next.mockResolvedValue(caeResponse);

    // No credential for CAE → authorizeRequestOnCaeChallenge returns false
    const policy = bearerTokenAuthenticationPolicy({
      scopes,
      credential: undefined,
    });

    const response = await policy.sendRequest(request, next);
    // Should return the 401 without retrying
    assert.equal(response.status, 401);
    expect(next).toHaveBeenCalledOnce();
  });

  it("does not retry when custom authorizeRequestOnChallenge returns false", async function () {
    // Covers line 279: shouldSendRequest = false from custom challenge callback
    const tokenExpiration = Date.now() + 1000 * 60;
    const credential: TokenCredential = {
      getToken: async () => ({
        token: "token",
        expiresOnTimestamp: tokenExpiration,
      }),
    };
    const scopes = ["test-scope"];
    const request = createPipelineRequest({ url: "https://example.com" });

    // Non-CAE 401 challenge
    const challengeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer authorization_uri="https://login.windows.net/", error="invalid_token"`,
      }),
      request,
      status: 401,
    };

    const next = vi.fn<SendRequest>();
    next.mockResolvedValue(challengeResponse);

    const policy = bearerTokenAuthenticationPolicy({
      scopes,
      credential,
      challengeCallbacks: {
        authorizeRequestOnChallenge: async () => false,
      },
    });

    const response = await policy.sendRequest(request, next);
    assert.equal(response.status, 401);
    expect(next).toHaveBeenCalledOnce();
  });

  it("handles CAE challenge when scopes is a single string (not array)", async function () {
    // Covers the `Array.isArray(scopes) ? scopes : [scopes]` false branches at lines 256, 271, 299
    const tokenExpiration = Date.now() + 1000 * 60;
    const credential: TokenCredential = {
      getToken: async () => ({
        token: "token",
        expiresOnTimestamp: tokenExpiration,
      }),
    };
    const request = createPipelineRequest({ url: "https://example.com" });

    // CAE challenge with valid claims
    const caeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer error="insufficient_claims", claims="eyJ0ZXN0IjoidmFsdWUifQ=="`,
      }),
      request,
      status: 401,
    };

    const successResponse: PipelineResponse = {
      headers: createHttpHeaders(),
      request,
      status: 200,
    };

    const next = vi.fn<SendRequest>();
    next.mockResolvedValueOnce(caeResponse).mockResolvedValueOnce(successResponse);

    const policy = bearerTokenAuthenticationPolicy({
      scopes: "single-scope-string",
      credential,
    });

    const response = await policy.sendRequest(request, next);
    assert.equal(response.status, 200);
    expect(next).toHaveBeenCalledTimes(2);
  });

  it("handles custom challenge then CAE challenge when scopes is a string", async function () {
    // Covers line 271 and 299 ternary branches when scopes is a string
    const tokenExpiration = Date.now() + 1000 * 60;
    const credential: TokenCredential = {
      getToken: async () => ({
        token: "token",
        expiresOnTimestamp: tokenExpiration,
      }),
    };
    const request = createPipelineRequest({ url: "https://example.com" });

    // Non-CAE challenge
    const nonCaeChallengeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer authorization_uri="https://login.windows.net/", error="invalid_token"`,
      }),
      request,
      status: 401,
    };

    // After custom handler: valid CAE challenge
    const caeChallengeResponse: PipelineResponse = {
      headers: createHttpHeaders({
        "WWW-Authenticate": `Bearer error="insufficient_claims", claims="eyJ0ZXN0IjoidmFsdWUifQ=="`,
      }),
      request,
      status: 401,
    };

    const successResponse: PipelineResponse = {
      headers: createHttpHeaders(),
      request,
      status: 200,
    };

    const next = vi.fn<SendRequest>();
    next
      .mockResolvedValueOnce(nonCaeChallengeResponse)
      .mockResolvedValueOnce(caeChallengeResponse)
      .mockResolvedValueOnce(successResponse);

    const policy = bearerTokenAuthenticationPolicy({
      scopes: "single-scope-string",
      credential,
      challengeCallbacks: {
        authorizeRequestOnChallenge: async (options) => {
          const token = await options.getAccessToken(["single-scope-string"], {});
          if (token) {
            options.request.headers.set("Authorization", `Bearer ${token.token}`);
            return true;
          }
          return false;
        },
      },
    });

    const response = await policy.sendRequest(request, next);
    assert.equal(response.status, 200);
    expect(next).toHaveBeenCalledTimes(3);
  });
});
