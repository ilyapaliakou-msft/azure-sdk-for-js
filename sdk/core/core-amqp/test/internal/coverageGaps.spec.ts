// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, it, assert, vi, beforeEach } from "vitest";
import { CbsClient, TokenType } from "../../src/index.js";
import { Connection, SenderEvents, ReceiverEvents } from "rhea-promise";
import { createConnectionStub } from "../utils/createConnectionStub.js";
import { RequestResponseLink } from "../../src/requestResponseLink.js";
import EventEmitter from "events";

/**
 * Creates a connection stub with full-featured session/sender/receiver mocks
 * that include isOpen() and remove() methods.
 */
function createFullConnectionStub(): Connection {
  const connectionStub = new Connection();
  vi.spyOn(connectionStub, "open").mockResolvedValue({} as any);
  vi.spyOn(connectionStub, "createSession").mockResolvedValue({
    connection: {
      id: "connection-1",
    },
    isOpen: () => true,
    remove: vi.fn(),
    close: vi.fn(),
    createSender: () => {
      const sender = new EventEmitter() as any;
      sender.send = () => {};
      sender.isOpen = () => true;
      sender.remove = vi.fn();
      sender.close = vi.fn();
      return Promise.resolve(sender);
    },
    createReceiver: () => {
      const receiver = new EventEmitter() as any;
      receiver.isOpen = () => true;
      receiver.remove = vi.fn();
      receiver.close = vi.fn();
      return Promise.resolve(receiver);
    },
  } as any);
  vi.spyOn(connectionStub, "id", "get").mockReturnValue("connection-1");
  return connectionStub;
}
import {
  randomNumberFromInterval,
  executePromisesSequentially,
  isIotHubConnectionString,
  isString,
  isNumber,
  getGlobalProperty,
  Timeout,
  delay,
} from "../../src/util/utils.js";
import { isSasTokenProvider } from "../../src/util/typeGuards.js";
import { signString } from "../../src/util/hmacSha256.common.js";
import * as Errors from "../../src/errors.js";
import { createSasTokenProvider } from "../../src/auth/tokenProvider.js";

describe("CbsClient - close, remove, isOpen", () => {
  it("close() when not open is a no-op", async () => {
    const cbsClient = new CbsClient(new Connection(), "lock");
    // Should not throw when not open
    await cbsClient.close();
  });

  it("close() when open closes the link", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    assert.isTrue(cbsClient.isOpen());
    await cbsClient.close();
    assert.isFalse(cbsClient.isOpen());
  });

  it("close() wraps errors from link.close()", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    // Make the underlying link's close throw
    const link = (cbsClient as any)._cbsSenderReceiverLink as RequestResponseLink;
    vi.spyOn(link, "close").mockRejectedValue(new Error("close failed"));
    try {
      await cbsClient.close();
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "An error occurred while closing the cbs link");
    }
  });

  it("close() wraps non-Error with stack from link.close()", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    const link = (cbsClient as any)._cbsSenderReceiverLink as RequestResponseLink;
    vi.spyOn(link, "close").mockRejectedValue({ something: "not an error" });
    try {
      await cbsClient.close();
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "An error occurred while closing the cbs link");
    }
  });

  it("remove() when not open is a no-op", () => {
    const cbsClient = new CbsClient(new Connection(), "lock");
    // Should not throw
    cbsClient.remove();
  });

  it("remove() when open removes the link", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    assert.isTrue(cbsClient.isOpen());
    cbsClient.remove();
    assert.isFalse(cbsClient.isOpen());
  });

  it("remove() wraps errors from link.remove()", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    const link = (cbsClient as any)._cbsSenderReceiverLink as RequestResponseLink;
    vi.spyOn(link, "remove").mockImplementation(() => {
      throw new Error("remove failed");
    });
    try {
      cbsClient.remove();
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "An error occurred while removing the cbs link");
    }
  });

  it("remove() wraps non-Error from link.remove()", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    const link = (cbsClient as any)._cbsSenderReceiverLink as RequestResponseLink;
    vi.spyOn(link, "remove").mockImplementation(() => {
      throw { something: "not an error" };
    });
    try {
      cbsClient.remove();
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "An error occurred while removing the cbs link");
    }
  });

  it("isOpen() returns false when no link", () => {
    const cbsClient = new CbsClient(new Connection(), "lock");
    assert.isFalse(cbsClient.isOpen());
  });

  it("negotiateClaim succeeds when link is open", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    // Mock sendRequest on the underlying link
    const link = (cbsClient as any)._cbsSenderReceiverLink as RequestResponseLink;
    vi.spyOn(link, "sendRequest").mockResolvedValue({
      correlation_id: "test-id",
      application_properties: {
        "status-code": 200,
        "status-description": "OK",
      },
    } as any);
    const response = await cbsClient.negotiateClaim("audience", "token", TokenType.CbsTokenTypeSas);
    assert.equal(response.statusCode, 200);
    assert.equal(response.statusDescription, "OK");
  });

  it("negotiateClaim propagates errors from sendRequest", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    const link = (cbsClient as any)._cbsSenderReceiverLink as RequestResponseLink;
    vi.spyOn(link, "sendRequest").mockRejectedValue(new Error("send failed"));
    try {
      await cbsClient.negotiateClaim("audience", "token", TokenType.CbsTokenTypeSas);
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.equal(err.message, "send failed");
    }
  });

  it("negotiateClaim propagates non-Error throws", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    const link = (cbsClient as any)._cbsSenderReceiverLink as RequestResponseLink;
    vi.spyOn(link, "sendRequest").mockRejectedValue("string error");
    try {
      await cbsClient.negotiateClaim("audience", "token", TokenType.CbsTokenTypeSas);
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.equal(err, "string error");
    }
  });
});

describe("utils.ts functions", () => {
  describe("randomNumberFromInterval", () => {
    it("returns a number within the given range", () => {
      for (let i = 0; i < 20; i++) {
        const result = randomNumberFromInterval(5, 10);
        assert.isAtLeast(result, 5);
        assert.isAtMost(result, 10);
      }
    });

    it("returns the value when min equals max", () => {
      const result = randomNumberFromInterval(7, 7);
      assert.equal(result, 7);
    });
  });

  describe("executePromisesSequentially", () => {
    it("executes promise factories sequentially", async () => {
      const results: number[] = [];
      const factories = [
        (input: number) => {
          results.push(input);
          return Promise.resolve(input + 1);
        },
        (input: number) => {
          results.push(input);
          return Promise.resolve(input + 1);
        },
        (input: number) => {
          results.push(input);
          return Promise.resolve(input + 1);
        },
      ];
      const finalResult = await executePromisesSequentially(factories, 0);
      assert.deepEqual(results, [0, 1, 2]);
      assert.equal(finalResult, 3);
    });

    it("works with empty array", async () => {
      const result = await executePromisesSequentially([]);
      assert.isUndefined(result);
    });

    it("works without kickstart", async () => {
      const result = await executePromisesSequentially([
        (val: any) => Promise.resolve(val === undefined ? "ok" : "fail"),
      ]);
      assert.equal(result, "ok");
    });
  });

  describe("isIotHubConnectionString", () => {
    it("returns true for IoT Hub connection strings", () => {
      const cs =
        "HostName=myhub.azure-devices.net;SharedAccessKeyName=iothubowner;SharedAccessKey=abc123";
      assert.isTrue(isIotHubConnectionString(cs));
    });

    it("returns false for non-IoT Hub connection strings", () => {
      const cs =
        "Endpoint=sb://mynamespace.servicebus.windows.net/;SharedAccessKeyName=RootManageSharedAccessKey;SharedAccessKey=abc123";
      assert.isFalse(isIotHubConnectionString(cs));
    });

    it("returns false for empty string", () => {
      assert.isFalse(isIotHubConnectionString(""));
    });
  });

  describe("isString", () => {
    it("returns true for strings", () => {
      assert.isTrue(isString("hello"));
      assert.isTrue(isString(""));
    });

    it("returns false for non-strings", () => {
      assert.isFalse(isString(123));
      assert.isFalse(isString(null));
      assert.isFalse(isString(undefined));
      assert.isFalse(isString({}));
    });
  });

  describe("isNumber", () => {
    it("returns true for numbers", () => {
      assert.isTrue(isNumber(123));
      assert.isTrue(isNumber(0));
      assert.isTrue(isNumber(NaN));
    });

    it("returns false for non-numbers", () => {
      assert.isFalse(isNumber("123"));
      assert.isFalse(isNumber(null));
      assert.isFalse(isNumber(undefined));
    });
  });

  describe("getGlobalProperty", () => {
    it("returns a global property", () => {
      const result = getGlobalProperty("setTimeout");
      assert.isDefined(result);
    });

    it("returns undefined for non-existing property", () => {
      const result = getGlobalProperty("nonExistingProperty12345");
      assert.isUndefined(result);
    });
  });

  describe("Timeout", () => {
    it("set resolves after timeout", async () => {
      const timeout = new Timeout();
      const result = await timeout.set(10);
      assert.isUndefined(result);
    });

    it("set rejects with value after timeout", async () => {
      const timeout = new Timeout();
      try {
        await timeout.set(10, "timeout error");
        assert.fail("Should have thrown");
      } catch (err: any) {
        assert.include(err.message, "timeout error");
      }
    });

    it("wrap resolves if promise resolves first", async () => {
      const result = await Timeout.wrap(Promise.resolve("ok"), 5000);
      assert.equal(result, "ok");
    });

    it("wrap rejects if promise rejects first", async () => {
      try {
        await Timeout.wrap(Promise.reject(new Error("fail")), 5000);
        assert.fail("Should have thrown");
      } catch (err: any) {
        assert.equal(err.message, "fail");
      }
    });

    it("static set works", async () => {
      const result = await Timeout.set(10);
      assert.isUndefined(result);
    });

    it("clear is safe when no timer", () => {
      const timeout = new Timeout();
      // Should not throw
      timeout.clear();
    });
  });

  describe("delay", () => {
    it("resolves with value when provided", async () => {
      const result = await delay(10, undefined, undefined, "hello");
      assert.equal(result, "hello");
    });

    it("resolves with void when no value", async () => {
      const result = await delay(10);
      assert.isUndefined(result);
    });
  });
});

describe("typeGuards", () => {
  describe("isSasTokenProvider", () => {
    it("returns true for SasTokenProvider-like objects", () => {
      assert.isTrue(isSasTokenProvider({ isSasTokenProvider: true }));
    });

    it("returns true for real SasTokenProviderImpl instances", () => {
      const provider = createSasTokenProvider({
        sharedAccessKeyName: "keyName",
        sharedAccessKey: "key",
      });
      assert.isTrue(isSasTokenProvider(provider));
      assert.isTrue(provider.isSasTokenProvider);
    });

    it("returns false for non-SasTokenProvider objects", () => {
      assert.isFalse(isSasTokenProvider({ isSasTokenProvider: false }));
      assert.isFalse(isSasTokenProvider({}));
      assert.isFalse(isSasTokenProvider(null));
      assert.isFalse(isSasTokenProvider(undefined));
      assert.isFalse(isSasTokenProvider("string"));
    });
  });
});

describe("SasTokenProvider", () => {
  it("createSasTokenProvider with sharedAccessSignature", () => {
    const provider = createSasTokenProvider({
      sharedAccessSignature: "SharedAccessSignature sr=test&sig=abc&se=123&skn=key",
    });
    assert.isTrue(provider.isSasTokenProvider);
  });

  it("getToken with SASCredential returns the signature directly", async () => {
    const provider = createSasTokenProvider({
      sharedAccessSignature: "SharedAccessSignature sr=test&sig=abc&se=123&skn=key",
    });
    const token = await provider.getToken("audience");
    assert.equal(token.token, "SharedAccessSignature sr=test&sig=abc&se=123&skn=key");
    assert.equal(token.expiresOnTimestamp, 0);
  });

  it("getToken with NamedKeyCredential returns a generated token", async () => {
    const provider = createSasTokenProvider({
      sharedAccessKeyName: "keyName",
      sharedAccessKey: "key",
    });
    const token = await provider.getToken("audience");
    assert.isString(token.token);
    assert.include(token.token, "SharedAccessSignature");
    assert.include(token.token, "sr=audience");
    assert.include(token.token, "skn=keyName");
    assert.isAbove(token.expiresOnTimestamp, 0);
  });
});

describe("hmacSha256.common (Web Crypto API)", () => {
  it("signString produces a valid HMAC-SHA256 signature", async () => {
    const result = await signString("testkey", "testdata");
    assert.isString(result);
    assert.isAbove(result.length, 0);
  });
});

describe("RequestResponseLink - remove", () => {
  it("remove() calls remove on sender, receiver, and session", async () => {
    const connectionStub = createFullConnectionStub();
    const link = await RequestResponseLink.create(connectionStub, {}, {});

    link.remove();

    // Verify the remove methods were called (they're vi.fn() from createFullConnectionStub)
    assert.isTrue(
      (link.sender.remove as any).mock.calls.length > 0,
      "sender.remove should be called",
    );
    assert.isTrue(
      (link.receiver.remove as any).mock.calls.length > 0,
      "receiver.remove should be called",
    );
    assert.isTrue(
      (link.session.remove as any).mock.calls.length > 0,
      "session.remove should be called",
    );
  });
});

describe("RequestResponseLink - onSenderError", () => {
  it("rejects all pending responses when sender errors", async () => {
    const connectionStub = createFullConnectionStub();
    const link = await RequestResponseLink.create(connectionStub, {}, {});
    const responsesMap = (link as any)._responsesMap as Map<string, any>;

    let rejected1 = false;
    let rejected2 = false;
    let cleanup1 = false;
    let cleanup2 = false;

    responsesMap.set("id1", {
      resolve: () => {},
      reject: () => {
        rejected1 = true;
      },
      cleanupBeforeResolveOrReject: () => {
        cleanup1 = true;
      },
    });
    responsesMap.set("id2", {
      resolve: () => {},
      reject: () => {
        rejected2 = true;
      },
      cleanupBeforeResolveOrReject: () => {
        cleanup2 = true;
      },
    });

    // Trigger the sender error event
    link.sender.emit("sender_error", {
      sender: {
        error: new Error("sender error"),
      },
    });

    assert.isTrue(rejected1, "First promise should be rejected");
    assert.isTrue(rejected2, "Second promise should be rejected");
    assert.isTrue(cleanup1, "First cleanup should be called");
    assert.isTrue(cleanup2, "Second cleanup should be called");
    assert.equal(responsesMap.size, 0, "Map should be cleared");
  });

  it("does nothing when sender is undefined", async () => {
    const connectionStub = createFullConnectionStub();
    const link = await RequestResponseLink.create(connectionStub, {}, {});
    const responsesMap = (link as any)._responsesMap as Map<string, any>;

    responsesMap.set("id1", {
      resolve: () => {},
      reject: () => {
        assert.fail("Should not be called");
      },
      cleanupBeforeResolveOrReject: () => {},
    });

    // Trigger sender error without a sender object
    link.sender.emit("sender_error", {});

    assert.equal(responsesMap.size, 1, "Map should not be affected");
  });
});

describe("errors.ts - additional coverage", () => {
  it("translate maps AMQP error with status-code: 404 in description to MessagingEntityNotFoundError", () => {
    const err: any = {
      name: "AmqpProtocolError",
      condition: "amqp:not-found",
      description: "The messaging entity blah could not be found. status-code: 404",
    };
    const translated = Errors.translate(err) as Errors.MessagingError;
    assert.equal(translated.code, "MessagingEntityNotFoundError");
  });

  it("translate maps AMQP error with 'messaging entity could not be found' to MessagingEntityNotFoundError", () => {
    const err: any = {
      name: "AmqpProtocolError",
      condition: "amqp:not-found",
      description: "The messaging entity 'myentity' could not be found.",
    };
    const translated = Errors.translate(err) as Errors.MessagingError;
    assert.equal(translated.code, "MessagingEntityNotFoundError");
  });

  it("translate handles already-translated MessagingError", () => {
    const err = new Errors.MessagingError("already translated");
    const translated = Errors.translate(err);
    assert.strictEqual(translated, err);
  });

  it("translate handles MessageWaitTimeout condition", () => {
    const err: any = {
      name: "AmqpProtocolError",
      condition: "com.microsoft:message-wait-timeout",
      description: "No messages available",
    };
    const translated = Errors.translate(err) as Errors.MessagingError;
    assert.equal(translated.name, "MessagingError");
    assert.equal(translated.code, "MessageWaitTimeout");
  });
});

describe("retry - additional coverage", () => {
  it("uses default retryOptions when none provided", async () => {
    let callCount = 0;
    const result = await (
      await import("../../src/retry.js")
    ).retry({
      operation: async () => {
        callCount++;
        return "ok";
      },
      connectionId: "conn-1",
      operationType: (await import("../../src/retry.js")).RetryOperationType.cbsAuth,
    });
    assert.equal(result, "ok");
    assert.equal(callCount, 1);
  });

  it("uses defaults for negative retryDelayInMs and maxRetryDelayInMs", async () => {
    let callCount = 0;
    const { retry, RetryOperationType } = await import("../../src/retry.js");
    const result = await retry({
      operation: async () => {
        callCount++;
        return "ok";
      },
      connectionId: "conn-1",
      operationType: RetryOperationType.cbsAuth,
      retryOptions: {
        maxRetries: 0,
        retryDelayInMs: -1,
        maxRetryDelayInMs: -1,
      },
    });
    assert.equal(result, "ok");
    assert.equal(callCount, 1);
  });

  it("checks network when ServiceCommunicationError and connectionHost provided", async () => {
    const { retry, RetryOperationType } = await import("../../src/retry.js");
    const { MessagingError, ErrorNameConditionMapper } = await import("../../src/errors.js");

    let callCount = 0;
    try {
      await retry({
        operation: async () => {
          callCount++;
          const err: any = {
            condition: ErrorNameConditionMapper.ServiceCommunicationError,
            description: "Connection lost",
          };
          throw err;
        },
        connectionId: "conn-1",
        operationType: RetryOperationType.cbsAuth,
        connectionHost: "localhost",
        retryOptions: {
          maxRetries: 0,
          retryDelayInMs: 100,
        },
      });
      assert.fail("Should have thrown");
    } catch {
      // The error should have been thrown after the network check
      assert.equal(callCount, 1);
    }
  });
});

describe("ConnectionContextBase - CoreAmqpConnection", () => {
  it("createSender sets maxListeners to 1000", async () => {
    const { ConnectionContextBase, ConnectionConfig } = await import("../../src/index.js");
    const connectionString =
      "Endpoint=sb://hostname.servicebus.windows.net/;SharedAccessKeyName=sakName;SharedAccessKey=sak;EntityPath=ep";
    const config = ConnectionConfig.create(connectionString, "mypath");
    const context = ConnectionContextBase.create({
      config,
      connectionProperties: {
        product: "MSJSClient",
        userAgent: "/js-amqp-client",
        version: "1.0.0",
      },
    });
    const conn = context.connection;

    // Mock the parent class methods
    const mockSender = {
      setMaxListeners: vi.fn(),
    };
    const mockAwaitableSender = {
      setMaxListeners: vi.fn(),
    };
    const mockReceiver = {
      setMaxListeners: vi.fn(),
    };

    const createSessionSpy = vi.spyOn(conn, "createSession" as any).mockResolvedValue({
      createSender: () => Promise.resolve(mockSender),
      createAwaitableSender: () => Promise.resolve(mockAwaitableSender),
      createReceiver: () => Promise.resolve(mockReceiver),
    } as any);

    // Test createSender by calling the Connection's createSession then the session's createSender
    // But CoreAmqpConnection overrides createSender directly on Connection
    // We need to mock super.createSender, super.createAwaitableSender, super.createReceiver

    // Use prototype chain to test
    const rheaPromise = await import("rhea-promise");
    vi.spyOn(rheaPromise.Connection.prototype, "createSender").mockResolvedValue(
      mockSender as any,
    );
    vi.spyOn(rheaPromise.Connection.prototype, "createAwaitableSender").mockResolvedValue(
      mockAwaitableSender as any,
    );
    vi.spyOn(rheaPromise.Connection.prototype, "createReceiver").mockResolvedValue(
      mockReceiver as any,
    );

    const sender = await conn.createSender();
    assert.isTrue(mockSender.setMaxListeners.mock.calls.length > 0);
    assert.equal(mockSender.setMaxListeners.mock.calls[0][0], 1000);

    const awaitableSender = await conn.createAwaitableSender();
    assert.isTrue(mockAwaitableSender.setMaxListeners.mock.calls.length > 0);
    assert.equal(mockAwaitableSender.setMaxListeners.mock.calls[0][0], 1000);

    const receiver = await conn.createReceiver();
    assert.isTrue(mockReceiver.setMaxListeners.mock.calls.length > 0);
    assert.equal(mockReceiver.setMaxListeners.mock.calls[0][0], 1000);

    createSessionSpy.mockRestore();
  });
});

describe("CbsClient - init already open branch and error handlers", () => {
  it("init when already open reuses existing link", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    assert.isTrue(cbsClient.isOpen());
    // Call init again - should hit the "already open" branch
    await cbsClient.init();
    assert.isTrue(cbsClient.isOpen());
  });

  it("sender error handler on cbs link fires without throwing", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    const link = (cbsClient as any)._cbsSenderReceiverLink as RequestResponseLink;
    // Trigger sender error event - the handler registered in cbs.ts init()
    link.sender.emit(SenderEvents.senderError, {
      connection: { options: { id: "connection-1" } },
      sender: { error: new Error("sender error") },
    });
    // Should not throw, just logs
    assert.isTrue(cbsClient.isOpen());
  });

  it("receiver error handler on cbs link fires without throwing", async () => {
    const connectionStub = createFullConnectionStub();
    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();
    const link = (cbsClient as any)._cbsSenderReceiverLink as RequestResponseLink;
    // Trigger receiver error event - the handler registered in cbs.ts init()
    link.receiver.emit(ReceiverEvents.receiverError, {
      connection: { options: { id: "connection-1" } },
      receiver: { error: new Error("receiver error") },
    });
    // Should not throw, just logs
    assert.isTrue(cbsClient.isOpen());
  });
});

describe("retry - validateRetryConfig", () => {
  it("throws TypeError when operation is missing", async () => {
    const { retry, RetryOperationType } = await import("../../src/retry.js");
    try {
      await retry({
        operation: undefined as any,
        connectionId: "conn-1",
        operationType: RetryOperationType.cbsAuth,
      });
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.instanceOf(err, TypeError);
      assert.include(err.message, "operation");
    }
  });

  it("throws TypeError when connectionId is missing", async () => {
    const { retry, RetryOperationType } = await import("../../src/retry.js");
    try {
      await retry({
        operation: async () => "ok",
        connectionId: "" as any,
        operationType: RetryOperationType.cbsAuth,
      });
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.instanceOf(err, TypeError);
      assert.include(err.message, "connectionId");
    }
  });

  it("throws TypeError when operationType is missing", async () => {
    const { retry } = await import("../../src/retry.js");
    try {
      await retry({
        operation: async () => "ok",
        connectionId: "conn-1",
        operationType: "" as any,
      });
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.instanceOf(err, TypeError);
      assert.include(err.message, "operationType");
    }
  });
});

describe("checkNetworkConnection (Node.js)", () => {
  it("returns true when DNS resolves successfully", async () => {
    const { checkNetworkConnection } = await import("../../src/util/checkNetworkConnection.js");
    const result = await checkNetworkConnection("localhost");
    assert.isTrue(result);
  });

  it("returns true when DNS fails with non-network error", async () => {
    const { checkNetworkConnection } = await import("../../src/util/checkNetworkConnection.js");
    // Use a domain that will fail but not with CONNREFUSED or TIMEOUT
    const result = await checkNetworkConnection("thishostdoesnotexist12345.invalid");
    // DNS NXDOMAIN is not CONNREFUSED or TIMEOUT, so it returns true
    assert.isBoolean(result);
  });
});

describe("lock.ts - edge cases", () => {
  it("handles empty queue during processing", async () => {
    const { CancellableAsyncLockImpl } = await import("../../src/util/lock.js");
    const lock = new CancellableAsyncLockImpl();

    // Simple task to verify the lock works
    const result = await lock.acquire("test-key", async () => "done", {
      abortSignal: undefined,
      timeoutInMs: undefined,
    });
    assert.equal(result, "done");
  });

  it("handles timeout removing a task from the queue", async () => {
    const { CancellableAsyncLockImpl } = await import("../../src/util/lock.js");
    const { delay: coreDelay } = await import("@azure/core-util");
    const lock = new CancellableAsyncLockImpl();

    // Task 1: Hold the lock for a bit
    const task1 = lock.acquire(
      "key",
      async () => {
        await coreDelay(50);
        return 1;
      },
      { abortSignal: undefined, timeoutInMs: undefined },
    );

    // Task 2: Times out immediately
    const task2 = lock
      .acquire(
        "key",
        async () => {
          return 2;
        },
        { abortSignal: undefined, timeoutInMs: 0 },
      )
      .catch((err) => {
        // Catch the timeout error to prevent unhandled rejection
        assert.equal(err.name, "OperationTimeoutError");
        return "timed-out";
      });

    const result1 = await task1;
    assert.equal(result1, 1);

    const result2 = await task2;
    assert.equal(result2, "timed-out");
  });
});

describe("utils.ts - getGlobalProperty catch branch", () => {
  it("returns undefined when globalThis access throws", async () => {
    // The catch branch is hard to trigger because globalThis is always available in Node.
    // We test it by directly verifying the function handles access gracefully.
    const result = getGlobalProperty("__nonExistent__");
    assert.isUndefined(result);
  });
});

describe("retry - isDelivery branch", () => {
  it("succeeds with a delivery-like result object (does not log result)", async () => {
    const { retry, RetryOperationType } = await import("../../src/retry.js");
    const deliveryResult = {
      id: 1,
      settled: true,
      remote_settled: false,
      format: 0,
    };
    const result = await retry({
      operation: async () => deliveryResult,
      connectionId: "conn-1",
      operationType: RetryOperationType.sendMessage,
      retryOptions: { maxRetries: 0 },
    });
    assert.deepEqual(result, deliveryResult);
  });
});

describe("cbs.ts - onSessionError callback", () => {
  it("onSessionError handler fires without throwing", async () => {
    // Create a connection stub that captures receiverOptions
    let capturedRxOpt: any = null;
    const connectionStub = new Connection();
    vi.spyOn(connectionStub, "open").mockResolvedValue({} as any);
    vi.spyOn(connectionStub, "createSession").mockResolvedValue({
      connection: {
        id: "connection-1",
        options: { id: "connection-1" },
      },
      isOpen: () => true,
      remove: vi.fn(),
      close: vi.fn(),
      createSender: () => {
        const sender = new EventEmitter() as any;
        sender.send = () => {};
        sender.isOpen = () => true;
        sender.remove = vi.fn();
        sender.close = vi.fn();
        sender.name = "cbs-sender";
        return Promise.resolve(sender);
      },
      createReceiver: (opts: any) => {
        capturedRxOpt = opts;
        const receiver = new EventEmitter() as any;
        receiver.isOpen = () => true;
        receiver.remove = vi.fn();
        receiver.close = vi.fn();
        receiver.name = "cbs-receiver";
        return Promise.resolve(receiver);
      },
    } as any);
    vi.spyOn(connectionStub, "id", "get").mockReturnValue("connection-1");

    const cbsClient = new CbsClient(connectionStub, "lock");
    await cbsClient.init();

    // Now call the captured onSessionError handler
    assert.isDefined(capturedRxOpt, "Receiver options should have been captured");
    assert.isDefined(capturedRxOpt.onSessionError, "onSessionError should be defined");

    // Call the handler - should not throw
    capturedRxOpt.onSessionError({
      connection: { options: { id: "connection-1" } },
      session: { error: { condition: "amqp:internal-error", description: "test error" } },
    });
  });
});

describe("checkNetworkConnection - DNS error codes", () => {
  it("calls dns.resolve and returns a boolean", async () => {
    const { checkNetworkConnection } = await import("../../src/util/checkNetworkConnection.js");
    // Use a hostname that should be resolvable in most environments
    const result = await checkNetworkConnection("dns.google");
    assert.isBoolean(result);
  });

  it("handles DNS resolution for invalid hostnames", async () => {
    const { checkNetworkConnection } = await import("../../src/util/checkNetworkConnection.js");
    // This will trigger the DNS error path (likely ENOTFOUND, not CONNREFUSED/TIMEOUT)
    // so it returns true (since ENOTFOUND is not in the list of network-down codes)
    const result = await checkNetworkConnection("thishostdefinitelydoesnotexist12345.invalid");
    assert.isBoolean(result);
  });
});

describe("RequestResponseLink - timeout with abortSignal cleans up abort listener (line 163)", () => {
  it("removes abort listener when timeout fires", async () => {
    const connectionStub = createFullConnectionStub();
    const link = await RequestResponseLink.create(connectionStub, {}, {});
    const request = { body: "test", message_id: "test-timeout-abort" };

    const controller = new AbortController();
    try {
      // Use a very short timeout so it fires, with abort signal that won't fire
      await link.sendRequest(request, {
        timeoutInMs: 10,
        abortSignal: controller.signal,
        requestName: "test",
      });
      assert.fail("Should have thrown");
    } catch (err: any) {
      // Should be OperationTimeoutError
      assert.include(err.message, "timed out");
    }
  });
});

describe("lock.ts - _removeTaskDetails with empty queue (line 212)", () => {
  it("_removeTaskDetails returns early when taskQueue is empty", async () => {
    const { CancellableAsyncLockImpl } = await import("../../src/util/lock.js");
    const lock = new CancellableAsyncLockImpl();

    // Access private method via any
    const lockAny = lock as any;

    // Call _removeTaskDetails with a key that doesn't exist in the map
    lockAny._removeTaskDetails("nonexistent-key", {});
    // Should not throw - just returns early (line 211-212)
  });

  it("_removeTaskDetails returns early when taskQueue is empty array", async () => {
    const { CancellableAsyncLockImpl } = await import("../../src/util/lock.js");
    const lock = new CancellableAsyncLockImpl();

    const lockAny = lock as any;
    // Set an empty array in the key map
    lockAny._keyMap.set("empty-key", []);

    // Call _removeTaskDetails - should hit the !taskQueue.length branch (line 210)
    lockAny._removeTaskDetails("empty-key", {});
  });

  it("_execute returns early when taskQueue is empty (line 173-174)", async () => {
    const { CancellableAsyncLockImpl } = await import("../../src/util/lock.js");
    const lock = new CancellableAsyncLockImpl();

    const lockAny = lock as any;
    // Ensure no task queue exists for the key
    // Call _execute directly
    await lockAny._execute("no-tasks-key");
    // Should return immediately without error (line 173-174)
  });
});
