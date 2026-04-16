// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * This test file uses vi.mock to mock node:dns before checkNetworkConnection.ts imports it.
 * This is necessary because ESM modules don't allow vi.spyOn on module exports.
 */
import { describe, it, assert, vi } from "vitest";

const { mockResolve } = vi.hoisted(() => ({
  mockResolve: vi.fn(),
}));

vi.mock("node:dns", () => ({
  CONNREFUSED: "ECONNREFUSED",
  TIMEOUT: "ETIMEOUT",
  resolve: mockResolve,
}));

import { checkNetworkConnection } from "../../src/util/checkNetworkConnection.js";

describe("checkNetworkConnection - mocked DNS", () => {
  it("returns false when DNS fails with ECONNREFUSED", async () => {
    mockResolve.mockImplementation((_host: string, cb: (err: any) => void) => {
      cb({ code: "ECONNREFUSED" });
    });
    const result = await checkNetworkConnection("example.com");
    assert.isFalse(result);
  });

  it("returns false when DNS fails with ETIMEOUT", async () => {
    mockResolve.mockImplementation((_host: string, cb: (err: any) => void) => {
      cb({ code: "ETIMEOUT" });
    });
    const result = await checkNetworkConnection("example.com");
    assert.isFalse(result);
  });

  it("returns true when DNS fails with other error", async () => {
    mockResolve.mockImplementation((_host: string, cb: (err: any) => void) => {
      cb({ code: "ENOTFOUND" });
    });
    const result = await checkNetworkConnection("example.com");
    assert.isTrue(result);
  });

  it("returns true when DNS resolves successfully", async () => {
    mockResolve.mockImplementation((_host: string, cb: (err: any) => void) => {
      cb(null);
    });
    const result = await checkNetworkConnection("example.com");
    assert.isTrue(result);
  });
});
