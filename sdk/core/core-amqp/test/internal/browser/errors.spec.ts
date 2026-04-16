// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, it, assert } from "vitest";
import { translate, MessagingError } from "../../../src/errors.js";

describe("translate - isBrowserWebsocketError (browser)", function () {
  it("translates a WebSocket error event into a MessagingError", function () {
    const ws = new WebSocket("wss://localhost:1");
    const errorEvent = new Event("error");
    Object.defineProperty(errorEvent, "target", { value: ws, writable: false });

    const result = translate(errorEvent);

    assert.instanceOf(result, MessagingError);
    assert.equal((result as MessagingError).code, "ServiceCommunicationError");
    assert.isFalse((result as MessagingError).retryable);
    assert.include(result.message, "Websocket");

    ws.close();
  });

  it("does not treat a plain error as a browser websocket error", function () {
    const plainError = new Error("not a websocket error");
    const result = translate(plainError);

    // A plain Error should be returned as-is, not wrapped as ServiceCommunicationError
    assert.equal(result, plainError);
  });
});
