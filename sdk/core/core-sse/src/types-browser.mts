// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * `IncomingMessage` is not available in the browser or React Native.
 */
export type IncomingMessage = never;

/**
 * `NodeJSReadableStream` is not available in the browser or React Native.
 */
export type NodeJSReadableStream = never;

/**
 * No-op in browser/React Native — Node streams are not available.
 */
export function cancelNodeStream(_stream: never): void {
  // unreachable
}
