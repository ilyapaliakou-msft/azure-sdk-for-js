// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, it, assert, vi } from "vitest";
import {
  type PipelineResponse,
  type SendRequest,
  createHttpHeaders,
  createPipelineRequest,
  ndJsonPolicy,
  setClientRequestIdPolicy,
  userAgentPolicy,
} from "../../../src/index.js";

describe("ndJsonPolicy - branch coverage", function () {
  function createMockNext(): SendRequest {
    const next = vi.fn<SendRequest>();
    next.mockImplementation(async (request) => ({
      headers: createHttpHeaders(),
      request,
      status: 200,
    }));
    return next;
  }

  it("passes through when body is not a string", async function () {
    const policy = ndJsonPolicy();
    const request = createPipelineRequest({ url: "https://example.com" });
    request.body = undefined;
    const next = createMockNext();
    await policy.sendRequest(request, next);
    assert.isUndefined(request.body);
  });

  it("passes through when body is a string not starting with [", async function () {
    const policy = ndJsonPolicy();
    const request = createPipelineRequest({ url: "https://example.com" });
    request.body = '{"key": "value"}';
    const next = createMockNext();
    await policy.sendRequest(request, next);
    assert.equal(request.body, '{"key": "value"}');
  });

  it("passes through when body is a number", async function () {
    const policy = ndJsonPolicy();
    const request = createPipelineRequest({ url: "https://example.com" });
    request.body = 42 as any;
    const next = createMockNext();
    await policy.sendRequest(request, next);
    assert.equal(request.body, 42);
  });
});

describe("setClientRequestIdPolicy - branch coverage", function () {
  it("does not overwrite an existing header", async function () {
    const policy = setClientRequestIdPolicy();
    const request = createPipelineRequest({ url: "https://example.com" });
    request.headers.set("x-ms-client-request-id", "custom-id");
    const next = vi.fn<SendRequest>();
    next.mockImplementation(async (req) => ({
      headers: createHttpHeaders(),
      request: req,
      status: 200,
    }));
    await policy.sendRequest(request, next);
    assert.equal(request.headers.get("x-ms-client-request-id"), "custom-id");
  });
});

describe("userAgentPolicy - branch coverage", function () {
  it("does not overwrite an existing User-Agent header", async function () {
    const policy = userAgentPolicy();
    const request = createPipelineRequest({ url: "https://example.com" });
    request.headers.set("User-Agent", "custom-agent");
    const next = vi.fn<SendRequest>();
    next.mockImplementation(async (req) => ({
      headers: createHttpHeaders(),
      request: req,
      status: 200,
    }));
    await policy.sendRequest(request, next);
    assert.equal(request.headers.get("User-Agent"), "custom-agent");
  });

  it("sets User-Agent header when not present", async function () {
    const policy = userAgentPolicy({ userAgentPrefix: "my-prefix" });
    const request = createPipelineRequest({ url: "https://example.com" });
    const next = vi.fn<SendRequest>();
    next.mockImplementation(async (req) => ({
      headers: createHttpHeaders(),
      request: req,
      status: 200,
    }));
    await policy.sendRequest(request, next);
    const ua = request.headers.get("User-Agent");
    assert.isDefined(ua);
    assert.isTrue(ua!.startsWith("my-prefix"));
  });
});
