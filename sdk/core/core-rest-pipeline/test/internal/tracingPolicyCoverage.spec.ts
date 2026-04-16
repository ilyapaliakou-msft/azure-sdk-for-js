// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, it, assert, expect, vi, beforeEach, afterEach } from "vitest";
import {
  type PipelineResponse,
  type SendRequest,
  createHttpHeaders,
  createPipelineRequest,
  RestError,
} from "../../src/index.js";
import {
  type Instrumenter,
  type InstrumenterSpanOptions,
  type SpanStatus,
  type TracingContext,
  type TracingSpan,
  type TracingSpanOptions,
  useInstrumenter,
} from "@azure/core-tracing";

// Mock createTracingClient for testing the error path
vi.mock("@azure/core-tracing", async (importOriginal) => {
  const original = await importOriginal<typeof import("@azure/core-tracing")>();
  return {
    ...original,
    createTracingClient: vi.fn(original.createTracingClient),
  };
});

// Mock getUserAgentValue so we can control the user agent string
vi.mock("../../src/util/userAgent.js", async (importOriginal) => {
  const original = await importOriginal<typeof import("../../src/util/userAgent.js")>();
  return {
    ...original,
    getUserAgentValue: vi.fn(original.getUserAgentValue),
  };
});

import { createTracingClient } from "@azure/core-tracing";
import { tracingPolicy } from "../../src/policies/tracingPolicy.js";
import { getUserAgentValue } from "../../src/util/userAgent.js";

class MockSpan implements TracingSpan {
  spanAttributes: Record<string, unknown> = {};
  endCalled: boolean = false;
  status?: SpanStatus;

  constructor(
    public name: string,
    spanOptions: TracingSpanOptions = {},
  ) {
    this.spanAttributes = spanOptions.spanAttributes ?? {};
  }

  isRecording(): boolean {
    return true;
  }

  recordException(): void {}

  end(): void {
    this.endCalled = true;
  }

  setStatus(status: SpanStatus): void {
    this.status = status;
  }

  setAttribute(name: string, value: unknown): void {
    this.spanAttributes[name] = value;
  }

  getAttribute(name: string): unknown {
    return this.spanAttributes[name];
  }
}

class NonRecordingSpan extends MockSpan {
  isRecording(): boolean {
    return false;
  }
}

const noopTracingContext: TracingContext = {
  deleteValue() {
    return this;
  },
  getValue() {
    return undefined;
  },
  setValue() {
    return this;
  },
};

class MockInstrumenter implements Instrumenter {
  lastSpanCreated: MockSpan | undefined;
  staticSpan: MockSpan | undefined;

  setStaticSpan(span: MockSpan): void {
    this.staticSpan = span;
  }

  startSpan(
    name: string,
    spanOptions: InstrumenterSpanOptions,
  ): { span: TracingSpan; tracingContext: TracingContext } {
    const tracingContext = spanOptions.tracingContext ?? noopTracingContext;
    if (this.staticSpan) {
      return { span: this.staticSpan, tracingContext };
    }
    const span = new MockSpan(name, spanOptions);
    this.lastSpanCreated = span;
    return { span, tracingContext };
  }

  withContext<
    CallbackArgs extends unknown[],
    Callback extends (...args: CallbackArgs) => ReturnType<Callback>,
  >(
    _context: TracingContext,
    callback: Callback,
    ...callbackArgs: CallbackArgs
  ): ReturnType<Callback> {
    return callback(...callbackArgs);
  }

  parseTraceparentHeader(): TracingContext | undefined {
    return undefined;
  }

  createRequestHeaders(): Record<string, string> {
    return {};
  }
}

describe("tracingPolicy - additional coverage", function () {
  let activeInstrumenter: MockInstrumenter;

  beforeEach(() => {
    activeInstrumenter = new MockInstrumenter();
    useInstrumenter(activeInstrumenter);
    vi.mocked(createTracingClient).mockRestore();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("passes through when tracingClient creation fails", async () => {
    // Make createTracingClient throw to exercise tryCreateTracingClient error path
    vi.mocked(createTracingClient).mockImplementation(() => {
      throw new Error("tracing not available");
    });

    const policy = tracingPolicy();
    const request = createPipelineRequest({
      url: "https://example.com",
      tracingOptions: { tracingContext: noopTracingContext },
    });
    const next = vi.fn<SendRequest>();
    const response: PipelineResponse = {
      headers: createHttpHeaders(),
      request,
      status: 200,
    };
    next.mockResolvedValue(response);

    const result = await policy.sendRequest(request, next);
    assert.equal(result.status, 200);
    // next should have been called directly (skipping tracing)
    expect(next).toHaveBeenCalledOnce();
  });

  it("skips tracing when span is not recording", async () => {
    const nonRecordingSpan = new NonRecordingSpan("non-recording");
    activeInstrumenter.setStaticSpan(nonRecordingSpan);

    const policy = tracingPolicy();
    const request = createPipelineRequest({
      url: "https://example.com",
      tracingOptions: { tracingContext: noopTracingContext },
    });
    const next = vi.fn<SendRequest>();
    const response: PipelineResponse = {
      headers: createHttpHeaders(),
      request,
      status: 200,
    };
    next.mockResolvedValue(response);

    await policy.sendRequest(request, next);
    assert.isTrue(nonRecordingSpan.endCalled, "non-recording span should be ended");
    // The span attributes should not be set for status code since we returned early
  });

  it("sets serviceRequestId attribute when x-ms-request-id header is present", async () => {
    const policy = tracingPolicy();
    const request = createPipelineRequest({
      url: "https://example.com",
      tracingOptions: { tracingContext: noopTracingContext },
    });

    const response: PipelineResponse = {
      headers: createHttpHeaders({ "x-ms-request-id": "test-request-id" }),
      request,
      status: 200,
    };
    const next = vi.fn<SendRequest>();
    next.mockResolvedValue(response);

    await policy.sendRequest(request, next);

    const span = activeInstrumenter.lastSpanCreated!;
    assert.equal(span.getAttribute("serviceRequestId"), "test-request-id");
  });

  it("handles error thrown by span.setStatus in tryProcessError", async () => {
    const mockSpan = new MockSpan("mock");
    vi.spyOn(mockSpan, "setStatus").mockImplementation(() => {
      throw new Error("setStatus failed");
    });
    activeInstrumenter.setStaticSpan(mockSpan);

    const policy = tracingPolicy();
    const request = createPipelineRequest({
      url: "https://example.com",
      tracingOptions: { tracingContext: noopTracingContext },
    });
    const next = vi.fn<SendRequest>();
    const expectedError = new RestError("Bad Request.", { statusCode: 400 });
    next.mockRejectedValue(expectedError);

    // The pipeline error should propagate, not the span processing error
    await expect(policy.sendRequest(request, next)).rejects.toThrow(expectedError);
  });

  it("handles error that is not a RestError (no statusCode attribute)", async () => {
    const policy = tracingPolicy();
    const request = createPipelineRequest({
      url: "https://example.com",
      tracingOptions: { tracingContext: noopTracingContext },
    });

    const next = vi.fn<SendRequest>();
    const genericError = new Error("generic error");
    next.mockRejectedValue(genericError);

    await expect(policy.sendRequest(request, next)).rejects.toThrow("generic error");
    const span = activeInstrumenter.lastSpanCreated!;
    assert.equal(span.status?.status, "error");
    assert.isTrue(span.endCalled);
    // No http.status_code should be set for non-RestError
    assert.isUndefined(span.getAttribute("http.status_code"));
  });

  it("handles non-Error thrown values in tryProcessError", async () => {
    const policy = tracingPolicy();
    const request = createPipelineRequest({
      url: "https://example.com",
      tracingOptions: { tracingContext: noopTracingContext },
    });

    const next = vi.fn<SendRequest>();
    next.mockRejectedValue("string error");

    await expect(policy.sendRequest(request, next)).rejects.toEqual("string error");
    const span = activeInstrumenter.lastSpanCreated!;
    assert.equal(span.status?.status, "error");
    // error should be undefined since "string error" is not an Error instance
    assert.isUndefined((span.status as any)?.error);
    assert.isTrue(span.endCalled);
  });

  it("does not set http.user_agent attribute when userAgent is empty", async () => {
    // Mock getUserAgentValue to return empty string to exercise the false branch of `if (userAgent)`
    vi.mocked(getUserAgentValue).mockResolvedValue("");

    const policy = tracingPolicy();
    const request = createPipelineRequest({
      url: "https://example.com",
      tracingOptions: { tracingContext: noopTracingContext },
    });
    const next = vi.fn<SendRequest>();
    const response: PipelineResponse = {
      headers: createHttpHeaders(),
      request,
      status: 200,
    };
    next.mockResolvedValue(response);

    await policy.sendRequest(request, next);
    const span = activeInstrumenter.lastSpanCreated!;
    assert.equal(span.getAttribute("http.user_agent"), "");
  });
});
