// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { IncomingMessage as NodeIncomingMessage } from "node:http";

export type { IncomingMessage } from "node:http";

/**
 * A Node.js Readable stream that also has a `destroy` method.
 */
export interface NodeJSReadableStream extends NodeJS.ReadableStream {
  /**
   * Destroy the stream. Optionally emit an 'error' event, and emit a
   * 'close' event (unless emitClose is set to false). After this call,
   * internal resources will be released.
   */
  destroy(error?: Error): void;
}

/**
 * Cancels a Node.js stream by ending its socket or destroying it.
 */
export function cancelNodeStream(stream: NodeIncomingMessage | NodeJSReadableStream): void {
  if ("socket" in stream && stream.socket) {
    stream.socket.end();
  } else if ("destroy" in stream) {
    stream.destroy();
  }
}
