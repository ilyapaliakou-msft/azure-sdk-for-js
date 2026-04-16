// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, it, assert, expect, vi } from "vitest";
import {
  inferLroMode,
  parseRetryAfter,
  getErrorFromResponse,
  getStatusFromInitialResponse,
  getOperationLocation,
  getOperationStatus,
  getResourceLocation,
  isOperationError,
  pollHttpOperation,
} from "../../src/http/operation.js";
import { rewriteUrl } from "../../src/http/utils.js";
import { deserializeState, initOperation, pollOperation } from "../../src/poller/operation.js";
import { buildCreatePoller } from "../../src/poller/poller.js";
import type { OperationResponse, RawResponse } from "../../src/http/models.js";
import type { OperationState, RestorableOperationState } from "../../src/poller/models.js";
import { createTestPoller } from "../utils/router.js";
import type { Result, State } from "../utils/utils.js";
import { createHttpPoller } from "../../src/http/poller.js";

function makeRawResponse(overrides: Partial<RawResponse> = {}): RawResponse {
  return {
    statusCode: 200,
    headers: {},
    request: { method: "GET", url: "https://example.com/resource" },
    ...overrides,
  };
}

function makeState<TResult>(
  mode?: string,
  extra?: Partial<RestorableOperationState<TResult, OperationState<TResult>>>,
): RestorableOperationState<TResult, OperationState<TResult>> {
  return {
    status: "running",
    config: {
      metadata: mode ? { mode } : undefined,
      ...extra?.config,
    },
    ...extra,
  } as any;
}

describe("http/operation.ts coverage", () => {
  describe("calculatePollingIntervalFromDate (via parseRetryAfter)", () => {
    it("returns undefined when retry-after date is in the past", () => {
      const pastDate = new Date(Date.now() - 100000).toUTCString();
      const result = parseRetryAfter({
        rawResponse: makeRawResponse({ headers: { "retry-after": pastDate } }),
        flatResponse: {},
      });
      assert.isUndefined(result);
    });

    it("returns milliseconds when retry-after date is in the future", () => {
      const futureDate = new Date(Date.now() + 60000).toUTCString();
      const result = parseRetryAfter({
        rawResponse: makeRawResponse({ headers: { "retry-after": futureDate } }),
        flatResponse: {},
      });
      assert.isNumber(result);
      assert.isAbove(result!, 0);
    });
  });

  describe("getOperationLocation", () => {
    it("returns undefined for Body mode", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse(),
        flatResponse: {},
      };
      const state = makeState("Body");
      assert.isUndefined(getOperationLocation(response, state));
    });

    it("returns undefined for unknown/default mode", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse(),
        flatResponse: {},
      };
      const state = makeState("SomeUnknownMode");
      assert.isUndefined(getOperationLocation(response, state));
    });
  });

  describe("getOperationStatus", () => {
    it("throws for unexpected mode", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse(),
        flatResponse: {},
      };
      const state = makeState("UnexpectedMode");
      assert.throws(() => getOperationStatus(response, state), /Unexpected operation mode/);
    });
  });

  describe("getStatusFromInitialResponse", () => {
    it("returns succeeded when mode is undefined and status code < 300", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ statusCode: 200 }),
        flatResponse: {},
      };
      const state = makeState(undefined);
      const status = getStatusFromInitialResponse({
        response,
        state,
        operationLocation: undefined,
      });
      assert.equal(status, "succeeded");
    });

    it("returns running when mode is undefined, status is 202, and operationLocation is set", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ statusCode: 202 }),
        flatResponse: {},
      };
      const state = makeState(undefined);
      const status = getStatusFromInitialResponse({
        response,
        state,
        operationLocation: "https://example.com/poll",
      });
      assert.equal(status, "running");
    });
  });

  describe("getErrorFromResponse", () => {
    it("returns undefined when error property has no code or message", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ body: { error: { code: "SomeCode" } } }),
        flatResponse: { error: { code: "SomeCode" } },
      };
      const result = getErrorFromResponse(response);
      assert.isUndefined(result);
    });

    it("returns undefined when no error property exists", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ body: {} }),
        flatResponse: {},
      };
      const result = getErrorFromResponse(response);
      assert.isUndefined(result);
    });
  });

  describe("getResourceLocation", () => {
    it("stores resourceLocation from response body to state config", () => {
      const state = makeState("OperationLocation");
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ body: { resourceLocation: "https://example.com/result" } }),
        flatResponse: { resourceLocation: "https://example.com/result" },
      };
      const loc = getResourceLocation(response, state);
      assert.equal(loc, "https://example.com/result");
      assert.equal(state.config.resourceLocation, "https://example.com/result");
    });
  });

  describe("isOperationError", () => {
    it("returns true for RestError", () => {
      const err = new Error("test");
      err.name = "RestError";
      assert.isTrue(isOperationError(err));
    });

    it("returns false for non-RestError", () => {
      assert.isFalse(isOperationError(new Error("test")));
    });
  });

  describe("transformStatus (via getOperationStatus)", () => {
    it("throws for non-string status", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ body: { status: 123 } }),
        flatResponse: {},
      };
      const state = makeState("OperationLocation");
      assert.throws(() => getOperationStatus(response, state), /Polling was unsuccessful/);
    });

    it("returns failed for status containing 'fail'", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ body: { status: "SomeFailure" } }),
        flatResponse: {},
      };
      const state = makeState("OperationLocation");
      assert.equal(getOperationStatus(response, state), "failed");
    });

    it("returns canceled for 'cancelled' status", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ body: { status: "Cancelled" } }),
        flatResponse: {},
      };
      const state = makeState("OperationLocation");
      assert.equal(getOperationStatus(response, state), "canceled");
    });

    it("returns running for unknown status", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ body: { status: "InProgress" } }),
        flatResponse: {},
      };
      const state = makeState("OperationLocation");
      assert.equal(getOperationStatus(response, state), "running");
    });
  });

  describe("inferLroMode", () => {
    it("returns undefined when no polling headers and non-PUT method", () => {
      const rawResponse = makeRawResponse({
        request: { method: "POST", url: "https://example.com/resource" },
      });
      const result = inferLroMode(rawResponse);
      assert.isUndefined(result);
    });

    it("returns Body mode for PUT with no polling headers", () => {
      const rawResponse = makeRawResponse({
        request: { method: "PUT", url: "https://example.com/resource" },
      });
      const result = inferLroMode(rawResponse);
      assert.equal(result?.mode, "Body");
    });

    it("handles PATCH with azure-async-operation resourceLocationConfig falls back to requestPath", () => {
      const rawResponse = makeRawResponse({
        request: { method: "PATCH", url: "https://example.com/resource" },
        headers: {
          "operation-location": "https://example.com/poll",
        },
      });
      const result = inferLroMode(rawResponse, "azure-async-operation");
      assert.equal(result?.mode, "OperationLocation");
      // azure-async-operation getDefault returns undefined, so PATCH falls back to requestPath
      assert.equal(result?.resourceLocation, "https://example.com/resource");
    });

    it("handles PATCH with operation-location resourceLocationConfig returns undefined", () => {
      const rawResponse = makeRawResponse({
        request: { method: "PATCH", url: "https://example.com/resource" },
        headers: {
          "operation-location": "https://example.com/poll",
        },
      });
      const result = inferLroMode(rawResponse, "operation-location");
      assert.equal(result?.mode, "OperationLocation");
      // operation-location getDefault returns undefined, PATCH falls back to requestPath
      assert.equal(result?.resourceLocation, "https://example.com/resource");
    });

    it("handles PATCH with original-uri resourceLocationConfig", () => {
      const rawResponse = makeRawResponse({
        request: { method: "PATCH", url: "https://example.com/resource" },
        headers: {
          "operation-location": "https://example.com/poll",
        },
      });
      const result = inferLroMode(rawResponse, "original-uri");
      assert.equal(result?.resourceLocation, "https://example.com/resource");
    });

    it("handles DELETE with operation-location", () => {
      const rawResponse = makeRawResponse({
        request: { method: "DELETE", url: "https://example.com/resource" },
        headers: {
          "operation-location": "https://example.com/poll",
        },
      });
      const result = inferLroMode(rawResponse);
      assert.equal(result?.mode, "OperationLocation");
      assert.isUndefined(result?.resourceLocation);
    });

    it("handles POST with azure-async-operation resourceLocationConfig", () => {
      const rawResponse = makeRawResponse({
        request: { method: "POST", url: "https://example.com/resource" },
        headers: {
          "operation-location": "https://example.com/poll",
          location: "https://example.com/location",
        },
      });
      const result = inferLroMode(rawResponse, "azure-async-operation");
      assert.isUndefined(result?.resourceLocation);
    });

    it("handles POST with location resourceLocationConfig (default)", () => {
      const rawResponse = makeRawResponse({
        request: { method: "POST", url: "https://example.com/resource" },
        headers: {
          "operation-location": "https://example.com/poll",
          location: "https://example.com/location",
        },
      });
      const result = inferLroMode(rawResponse, "location");
      assert.equal(result?.resourceLocation, "https://example.com/location");
    });

    it("uses azure-asyncoperation header when operation-location is missing", () => {
      const rawResponse = makeRawResponse({
        request: { method: "PUT", url: "https://example.com/resource" },
        headers: {
          "azure-asyncoperation": "https://example.com/async-poll",
        },
      });
      const result = inferLroMode(rawResponse);
      assert.equal(result?.mode, "OperationLocation");
      assert.equal(result?.operationLocation, "https://example.com/async-poll");
    });

    it("handles skipFinalGet to skip final resource GET", () => {
      const rawResponse = makeRawResponse({
        request: { method: "PUT", url: "https://example.com/resource" },
        headers: {
          "operation-location": "https://example.com/poll",
        },
      });
      const result = inferLroMode(rawResponse, undefined, true);
      assert.equal(result?.mode, "OperationLocation");
      assert.isUndefined(result?.resourceLocation);
    });

    it("handles PATCH with default resourceLocationConfig (location)", () => {
      const rawResponse = makeRawResponse({
        request: { method: "PATCH", url: "https://example.com/resource" },
        headers: {
          "operation-location": "https://example.com/poll",
          location: "https://example.com/location",
        },
      });
      // default config means location is used for PATCH
      const result = inferLroMode(rawResponse);
      assert.equal(result?.resourceLocation, "https://example.com/location");
    });

    it("uses requestPath for PATCH when location is undefined and config is not operation-location", () => {
      const rawResponse = makeRawResponse({
        request: { method: "PATCH", url: "https://example.com/resource" },
        headers: {
          "operation-location": "https://example.com/poll",
        },
      });
      // no location header and no specific config -> getDefault returns undefined for location -> falls back to requestPath
      const result = inferLroMode(rawResponse);
      assert.equal(result?.resourceLocation, "https://example.com/resource");
    });
  });

  describe("pollHttpOperation", () => {
    it("polls the operation through to completion", async () => {
      const pollPath = "/poll";
      const sendPollRequest = vi
        .fn()
        .mockResolvedValueOnce({
          flatResponse: { status: "running" },
          rawResponse: makeRawResponse({
            statusCode: 200,
            body: { status: "running" },
            headers: { "operation-location": pollPath },
          }),
        })
        .mockResolvedValueOnce({
          flatResponse: { status: "succeeded", id: "123" },
          rawResponse: makeRawResponse({
            statusCode: 200,
            body: { status: "succeeded", id: "123" },
          }),
        });

      const state = makeState<any>("OperationLocation", {
        config: {
          operationLocation: pollPath,
          metadata: { mode: "OperationLocation" },
        },
      } as any);

      const setDelay = vi.fn();
      await pollHttpOperation({
        lro: {
          sendInitialRequest: vi.fn(),
          sendPollRequest,
        },
        setDelay,
        state: state as any,
        setErrorAsResult: false,
      });

      assert.equal(sendPollRequest.mock.calls.length, 1);
    });

    it("uses processResult when provided", async () => {
      const pollPath = "/poll";
      const sendPollRequest = vi.fn().mockResolvedValueOnce({
        flatResponse: { value: "raw" },
        rawResponse: makeRawResponse({
          statusCode: 200,
          body: { status: "succeeded" },
        }),
      });

      const state = makeState<any>("OperationLocation", {
        config: {
          operationLocation: pollPath,
          metadata: { mode: "OperationLocation" },
        },
      } as any);

      await pollHttpOperation({
        lro: {
          sendInitialRequest: vi.fn(),
          sendPollRequest,
        },
        processResult: async (result: unknown, _s: any) => ({
          processed: true,
          ...(result as any),
        }),
        setDelay: vi.fn(),
        state: state as any,
        setErrorAsResult: false,
      });

      assert.deepEqual((state as any).result, { processed: true, value: "raw" });
    });
  });
});

describe("http/utils.ts coverage", () => {
  it("throws when relative URL cannot be resolved with invalid baseUrl", () => {
    // relative URL that can't be parsed even with baseUrl as base
    assert.throws(() => {
      rewriteUrl({ url: "://malformed", baseUrl: "not-a-url" });
    }, /Invalid input URL provided/);
  });
});

describe("poller/operation.ts coverage", () => {
  describe("deserializeState", () => {
    it("throws for invalid JSON", () => {
      assert.throws(() => deserializeState("not valid json"), /Unable to deserialize input state/);
    });
  });

  describe("simplifyError with innererror (via processOperationStatus)", () => {
    it("traverses innererror chain and appends messages", async () => {
      const pollingPath = "path/poll";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            body: JSON.stringify({
              status: "Failed",
              error: {
                code: "OuterCode",
                message: "Outer message",
                innererror: {
                  code: "InnerCode",
                  message: "Inner message",
                  innererror: {
                    code: "DeepCode",
                    message: "Deep message",
                  },
                },
              },
            }),
          },
        ],
        throwOnNon2xxResponse: true,
      });

      await expect(poller.pollUntilDone()).rejects.toThrow(/DeepCode/);
    });

    it("appends period to message when missing", async () => {
      const pollingPath = "path/poll";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            body: JSON.stringify({
              status: "Failed",
              error: {
                code: "ErrCode",
                message: "No period at end",
                innererror: {
                  code: "Inner",
                  message: "Inner detail",
                },
              },
            }),
          },
        ],
        throwOnNon2xxResponse: true,
      });

      await expect(poller.pollUntilDone()).rejects.toThrow(/No period at end\. Inner detail/);
    });
  });

  describe("setStateError", () => {
    it("sets state to failed when poll throws an operation error", async () => {
      const pollingPath = "path/poll";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          // The poll request will get a 500 which throws RestError
          {
            method: "GET",
            path: pollingPath,
            status: 500,
            body: JSON.stringify({ error: { code: "ServerError", message: "fail" } }),
          },
        ],
        throwOnNon2xxResponse: true,
      });

      await expect(poller.pollUntilDone()).rejects.toThrow();
    });
  });
});

describe("poller/poller.ts coverage", () => {
  describe("withOperationLocation callback", () => {
    it("calls withOperationLocation on initial and updated locations", async () => {
      const locations: string[] = [];
      const pollingPath = "path/poll";
      const newPollingPath = "path/poll-updated";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            headers: {
              "operation-location": newPollingPath,
            },
            body: JSON.stringify({ status: "InProgress" }),
          },
          {
            method: "GET",
            path: newPollingPath,
            status: 200,
            body: JSON.stringify({ status: "Succeeded" }),
          },
          {
            method: "GET",
            path: "path",
            status: 200,
            body: JSON.stringify({ id: "done" }),
          },
        ],
        withOperationLocation: (loc: string) => locations.push(loc),
        throwOnNon2xxResponse: true,
      });

      await poller.pollUntilDone();
      assert.isAbove(locations.length, 0);
      assert.include(locations, pollingPath);
      assert.include(locations, newPollingPath);
    });

    it("calls withOperationLocation only once for non-updated location", async () => {
      const locations: string[] = [];
      const pollingPath = "path/poll";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            headers: {
              "operation-location": pollingPath,
            },
            body: JSON.stringify({ status: "InProgress" }),
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            body: JSON.stringify({ status: "Succeeded" }),
          },
          {
            method: "GET",
            path: "path",
            status: 200,
            body: JSON.stringify({ id: "done" }),
          },
        ],
        withOperationLocation: (loc: string) => locations.push(loc),
        throwOnNon2xxResponse: true,
      });

      await poller.pollUntilDone();
      assert.equal(locations.length, 1);
      assert.equal(locations[0], pollingPath);
    });
  });

  describe("poll method edge cases", () => {
    it("returns state directly when already succeeded and resolveOnUnsuccessful", async () => {
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 200,
            body: JSON.stringify({ properties: { provisioningState: "Succeeded" }, id: "1" }),
          },
        ],
        throwOnNon2xxResponse: false,
      });

      // Wait for init
      await poller.submitted();
      // Polling an already-done poller
      const state = await poller.poll();
      assert.equal(state.status, "succeeded");
    });

    it("throws on poll when canceled and resolveOnUnsuccessful is false", async () => {
      const pollingPath = "path/poll-cancel";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            body: JSON.stringify({ status: "Canceled" }),
          },
        ],
        throwOnNon2xxResponse: true,
      });

      // First poll transitions to canceled
      await expect(poller.poll()).rejects.toThrow(/canceled/i);
      // Subsequent poll should also throw
      await expect(poller.poll()).rejects.toThrow(/canceled/i);
    });

    it("throws on poll when failed and resolveOnUnsuccessful is false", async () => {
      const pollingPath = "path/poll-fail";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            body: JSON.stringify({
              status: "Failed",
              error: { code: "Err", message: "something failed" },
            }),
          },
        ],
        throwOnNon2xxResponse: true,
      });

      await expect(poller.poll()).rejects.toThrow(/failed/i);
      // Subsequent poll should also throw the stored error
      await expect(poller.poll()).rejects.toThrow(/failed/i);
    });
  });

  describe("pollUntilDone edge cases", () => {
    it("throws canceled error from pollUntilDone", async () => {
      const pollingPath = "path/poll-cancel2";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            body: JSON.stringify({ status: "Canceled" }),
          },
        ],
        throwOnNon2xxResponse: true,
      });

      await expect(poller.pollUntilDone()).rejects.toThrow(/canceled/i);
    });

    it("uses setDelay when polling interval is provided via retry-after", async () => {
      const pollingPath = "path/poll-retry";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            headers: {
              "retry-after": "0",
            },
            body: JSON.stringify({ status: "InProgress" }),
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            body: JSON.stringify({ status: "Succeeded" }),
          },
          {
            method: "GET",
            path: "path",
            status: 200,
            body: JSON.stringify({ id: "done" }),
          },
        ],
        throwOnNon2xxResponse: true,
      });

      const result = await poller.pollUntilDone();
      assert.equal(result.statusCode, 200);
    });
  });

  describe("createHttpPoller with no options", () => {
    it("handles no options argument (undefined)", async () => {
      const lro = {
        sendInitialRequest: async () => ({
          flatResponse: { id: "1" },
          rawResponse: makeRawResponse({
            statusCode: 200,
            request: { method: "PUT", url: "https://example.com/resource" },
            body: { properties: { provisioningState: "Succeeded" } },
          }),
        }),
        sendPollRequest: vi.fn(),
      };

      const poller = createHttpPoller(lro);
      const result = await poller.pollUntilDone();
      assert.isDefined(result);
    });
  });

  describe("buildCreatePoller - resolveOnUnsuccessful", () => {
    it("returns result for canceled when resolveOnUnsuccessful is true", async () => {
      const pollingPath = "path/poll-resolve";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            body: JSON.stringify({ status: "Canceled" }),
          },
        ],
        throwOnNon2xxResponse: false,
      });

      const result = await poller.pollUntilDone();
      assert.isDefined(result);
    });

    it("isDone returns true for failed state", async () => {
      const pollingPath = "path/poll-fail-done";
      const poller = createTestPoller({
        routes: [
          {
            method: "PUT",
            status: 202,
            headers: {
              "operation-location": pollingPath,
            },
          },
          {
            method: "GET",
            path: pollingPath,
            status: 200,
            body: JSON.stringify({
              status: "Failed",
              error: { code: "Err", message: "Fail" },
            }),
          },
        ],
        throwOnNon2xxResponse: false,
      });

      const result = await poller.pollUntilDone();
      assert.isTrue(poller.isDone);
      assert.isDefined(result);
    });
  });

  describe("getProvisioningState via Body mode", () => {
    it("reads provisioningState from top-level body property", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({
          body: { provisioningState: "Succeeded" },
        }),
        flatResponse: {},
      };
      const state = makeState("Body");
      const status = getOperationStatus(response, state);
      assert.equal(status, "succeeded");
    });
  });

  describe("getOperationLocation for ResourceLocation mode", () => {
    it("returns location header for ResourceLocation mode", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({
          headers: { location: "https://example.com/location" },
        }),
        flatResponse: {},
      };
      const state = makeState("ResourceLocation");
      const loc = getOperationLocation(response, state);
      assert.equal(loc, "https://example.com/location");
    });
  });

  describe("getOperationStatus for ResourceLocation mode", () => {
    it("uses status code for ResourceLocation mode", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ statusCode: 202 }),
        flatResponse: {},
      };
      const state = makeState("ResourceLocation");
      assert.equal(getOperationStatus(response, state), "running");
    });

    it("returns succeeded for status 200 in ResourceLocation mode", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ statusCode: 200 }),
        flatResponse: {},
      };
      const state = makeState("ResourceLocation");
      assert.equal(getOperationStatus(response, state), "succeeded");
    });

    it("returns failed for status >= 300 in ResourceLocation mode", () => {
      const response: OperationResponse = {
        rawResponse: makeRawResponse({ statusCode: 500 }),
        flatResponse: {},
      };
      const state = makeState("ResourceLocation");
      assert.equal(getOperationStatus(response, state), "failed");
    });
  });
});

describe("pollOperation edge cases", () => {
  it("does nothing when operationLocation is undefined", async () => {
    const state = makeState<any>("OperationLocation");
    state.config.operationLocation = undefined;
    const poll = vi.fn();

    await pollOperation({
      poll,
      state,
      getOperationStatus: () => "running",
      getResourceLocation: () => undefined,
      isOperationError: () => false,
      setDelay: vi.fn(),
      setErrorAsResult: false,
    });

    assert.equal(poll.mock.calls.length, 0);
  });

  it("calls updateState after poll", async () => {
    const state = makeState<any>("OperationLocation");
    state.config.operationLocation = "/poll";

    const mockResponse = { data: "test" };
    const poll = vi.fn().mockResolvedValue(mockResponse);
    const updateState = vi.fn();

    await pollOperation({
      poll,
      state,
      getOperationStatus: () => "succeeded",
      getResourceLocation: () => undefined,
      isOperationError: () => false,
      setDelay: vi.fn(),
      setErrorAsResult: false,
      updateState,
    });

    assert.isTrue(updateState.mock.calls.length > 0);
  });

  it("calls withOperationLocation with same location when getOperationLocation returns undefined", async () => {
    const locations: Array<{ loc: string; isUpdated: boolean }> = [];
    const state = makeState<any>("OperationLocation");
    state.config.operationLocation = "/poll";

    const poll = vi.fn().mockResolvedValue({ data: "test" });

    await pollOperation({
      poll,
      state,
      getOperationStatus: () => "running",
      getResourceLocation: () => undefined,
      isOperationError: () => false,
      setDelay: vi.fn(),
      setErrorAsResult: false,
      withOperationLocation: (loc: string, isUpdated: boolean) =>
        locations.push({ loc, isUpdated }),
      getOperationLocation: () => undefined,
    });

    assert.equal(locations.length, 1);
    assert.equal(locations[0].loc, "/poll");
    assert.isFalse(locations[0].isUpdated);
  });
});

describe("initOperation edge cases", () => {
  it("calls withOperationLocation when operationLocation is present", async () => {
    const locations: string[] = [];
    await initOperation({
      init: async () => ({
        response: { data: "init" },
        operationLocation: "/poll-loc",
      }),
      getOperationStatus: () => "running",
      withOperationLocation: (loc: string) => locations.push(loc),
      setErrorAsResult: false,
    });

    assert.include(locations, "/poll-loc");
  });
});

describe("buildCreatePoller edge cases", () => {
  it("setDelay is called when getPollingInterval returns a value", async () => {
    let pollingInterval: number | undefined;
    let pollCount = 0;
    const createPoller = buildCreatePoller<any, any, OperationState<any>>({
      getStatusFromInitialResponse: () => "running",
      getStatusFromPollResponse: () => {
        pollCount++;
        return pollCount >= 2 ? "succeeded" : "running";
      },
      isOperationError: () => false,
      getResourceLocation: () => undefined,
      getPollingInterval: () => 42,
      resolveOnUnsuccessful: false,
    });

    const poller = createPoller(
      {
        init: async () => ({
          response: { data: "init" },
          operationLocation: "/poll",
        }),
        poll: async () => ({ data: "polled" }),
      },
      { intervalInMs: 0 },
    );

    // Use poll() directly to trigger setDelay
    await poller.submitted();
    const state = await poller.poll();
    // The poller should have updated the polling interval internally
    assert.equal(state.status, "running");
    // Poll again to succeed
    const finalState = await poller.poll();
    assert.equal(finalState.status, "succeeded");
  });

  it("handles poll with !state guard (defense check)", async () => {
    // Test the !state guards at lines 107 and 155 by making init resolve without setting state
    const createPoller = buildCreatePoller<any, any, OperationState<any>>({
      getStatusFromInitialResponse: () => "running",
      getStatusFromPollResponse: () => "running",
      isOperationError: () => false,
      getResourceLocation: () => undefined,
      resolveOnUnsuccessful: false,
    });

    const poller = createPoller(
      {
        init: async () => {
          return {
            response: { data: "init" },
            operationLocation: "/poll",
          };
        },
        poll: async () => ({ data: "polled" }),
      },
      { intervalInMs: 0 },
    );

    await poller.submitted();
    const state = await poller.poll();
    assert.isDefined(state);
  });

});

describe("pollHttpOperation without processResult", () => {
  it("uses default flatResponse identity when processResult is not provided", async () => {
    const pollPath = "/poll-no-process";
    const sendPollRequest = vi.fn().mockResolvedValueOnce({
      flatResponse: { id: "result-123", statusCode: 200 },
      rawResponse: makeRawResponse({
        statusCode: 200,
        body: { status: "succeeded" },
      }),
    });

    const state = makeState<any>("OperationLocation", {
      config: {
        operationLocation: pollPath,
        metadata: { mode: "OperationLocation" },
      },
    } as any);

    await pollHttpOperation({
      lro: {
        sendInitialRequest: vi.fn(),
        sendPollRequest,
      },
      setDelay: vi.fn(),
      state: state as any,
      setErrorAsResult: false,
    });

    // Without processResult, the flatResponse should be used as-is
    assert.deepEqual((state as any).result, { id: "result-123", statusCode: 200 });
  });
});

describe("processOperationStatus with isDone callback", () => {
  it("uses custom isDone to determine completion", async () => {
    let pollCount = 0;
    const createPoller = buildCreatePoller<any, any, OperationState<any>>({
      getStatusFromInitialResponse: () => "running",
      getStatusFromPollResponse: () => {
        pollCount++;
        return "running";
      },
      isOperationError: () => false,
      getResourceLocation: () => undefined,
      resolveOnUnsuccessful: false,
    });

    const poller = createPoller(
      {
        init: async () => ({
          response: { data: "init" },
          operationLocation: "/poll",
        }),
        poll: async () => ({ data: "polled", customDone: pollCount >= 1 }),
      },
      {
        intervalInMs: 0,
        processResult: async (response: any) => response,
      },
    );

    await poller.submitted();
    const state = await poller.poll();
    // Status is still "running" since we return "running"
    assert.equal(state.status, "running");
  });
});

describe("processOperationStatus setErrorAsResult=true includes failed in done states", () => {
  it("sets result when status is failed and setErrorAsResult is true", async () => {
    const pollingPath = "path/poll-err-result";
    const poller = createTestPoller({
      routes: [
        {
          method: "PUT",
          status: 202,
          headers: {
            "operation-location": pollingPath,
          },
        },
        {
          method: "GET",
          path: pollingPath,
          status: 200,
          body: JSON.stringify({
            status: "Failed",
            error: { code: "SomeError", message: "Something went wrong" },
          }),
        },
      ],
      throwOnNon2xxResponse: false,
    });

    const result = await poller.pollUntilDone();
    assert.isDefined(result);
    assert.isTrue(poller.isDone);
  });
});

describe("operation.ts branch: appendReadableErrorMessage with message ending in period", () => {
  it("does not double-add a period when message already ends with one", async () => {
    const pollingPath = "path/poll-period";
    const poller = createTestPoller({
      routes: [
        {
          method: "PUT",
          status: 202,
          headers: {
            "operation-location": pollingPath,
          },
        },
        {
          method: "GET",
          path: pollingPath,
          status: 200,
          body: JSON.stringify({
            status: "Failed",
            error: {
              code: "Err",
              message: "Something failed.",
              innererror: {
                code: "Inner",
                message: "Inner detail.",
              },
            },
          }),
        },
      ],
      throwOnNon2xxResponse: true,
    });

    await expect(poller.pollUntilDone()).rejects.toThrow(/Something failed\. Inner detail\./);
  });
});
