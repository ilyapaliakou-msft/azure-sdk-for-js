// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  type PipelineResponse,
  type SendRequest,
  auxiliaryAuthenticationHeaderPolicy,
  createHttpHeaders,
  createPipelineRequest,
} from "../../src/index.js";
import { describe, it, assert, vi, beforeEach, afterEach } from "vitest";

describe("AuxiliaryAuthenticationHeaderPolicy - additional coverage", function () {
  beforeEach(() => {
    vi.useFakeTimers({ now: Date.now() });
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  it("skips setting header when credentials is an empty array", async function () {
    const request = createPipelineRequest({ url: "https://example.com" });
    const successResponse: PipelineResponse = {
      headers: createHttpHeaders(),
      request,
      status: 200,
    };
    const next = vi.fn<SendRequest>();
    next.mockResolvedValue(successResponse);

    const policy = auxiliaryAuthenticationHeaderPolicy({
      scopes: ["scope1"],
      credentials: [],
    });

    await policy.sendRequest(request, next);
    assert.isUndefined(request.headers.get("x-ms-authorization-auxiliary"));
  });

  it("skips setting header when credentials is undefined", async function () {
    const request = createPipelineRequest({ url: "https://example.com" });
    const successResponse: PipelineResponse = {
      headers: createHttpHeaders(),
      request,
      status: 200,
    };
    const next = vi.fn<SendRequest>();
    next.mockResolvedValue(successResponse);

    const policy = auxiliaryAuthenticationHeaderPolicy({
      scopes: ["scope1"],
      credentials: undefined,
    });

    await policy.sendRequest(request, next);
    assert.isUndefined(request.headers.get("x-ms-authorization-auxiliary"));
  });
});
