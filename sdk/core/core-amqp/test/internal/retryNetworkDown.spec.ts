// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * This test file uses vi.mock to mock checkNetworkConnection before retry.ts imports it.
 * This is necessary because ESM modules don't allow vi.spyOn on module exports.
 */
import { describe, it, assert, vi } from "vitest";

const { mockCheckNetwork } = vi.hoisted(() => ({
  mockCheckNetwork: vi.fn().mockResolvedValue(false),
}));

vi.mock("../../src/util/checkNetworkConnection.js", () => ({
  checkNetworkConnection: mockCheckNetwork,
}));

import { retry, RetryOperationType } from "../../src/retry.js";
import { MessagingError } from "../../src/errors.js";

describe("retry - ConnectionLostError when checkNetworkConnection returns false", () => {
  it("marks ServiceCommunicationError as ConnectionLostError when network is down", async () => {
    let callCount = 0;
    try {
      await retry({
        operation: async () => {
          callCount++;
          // Create a MessagingError with name=ServiceCommunicationError and retryable=false
          // This matches the condition in retry.ts lines 234-237
          const err = new MessagingError("Connection lost");
          err.name = "ServiceCommunicationError";
          err.retryable = false;
          throw err;
        },
        connectionId: "conn-1",
        operationType: RetryOperationType.cbsAuth,
        connectionHost: "nonexistent.host.invalid",
        retryOptions: {
          maxRetries: 1,
          retryDelayInMs: 10,
        },
      });
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.isTrue(
        mockCheckNetwork.mock.calls.length > 0,
        "checkNetworkConnection should have been called",
      );
      assert.isAbove(callCount, 1, "Operation should have been retried");
    }
  });
});
