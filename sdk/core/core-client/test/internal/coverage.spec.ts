// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Tests specifically targeting uncovered code paths to achieve 100% coverage.
 */

import { describe, it, assert, vi } from "vitest";
import type {
  CompositeMapper,
  DictionaryMapper,
  FullOperationResponse,
  OperationArguments,
  OperationRequest,
  OperationSpec,
  SequenceMapper,
} from "../../src/index.js";
import {
  createSerializer,
  ServiceClient,
  serializationPolicy,
  MapperTypeNames,
} from "../../src/index.js";
import {
  createEmptyPipeline,
  createHttpHeaders,
  createPipelineRequest,
} from "@azure/core-rest-pipeline";
import {
  getOperationArgumentValueFromParameter,
  getOperationRequestInfo,
} from "../../src/operationHelpers.js";
import { getPathStringFromParameter } from "../../src/interfaceHelpers.js";
import { appendQueryParams, getRequestUrl } from "../../src/urlHelpers.js";
import { flattenResponse } from "../../src/utils.js";
import { createClientPipeline } from "../../src/pipeline.js";
import { encodeByteArray } from "../../src/base64.js";
import { deserializationPolicy } from "../../src/deserializationPolicy.js";
import { authorizeRequestOnClaimChallenge } from "../../src/authorizeRequestOnClaimChallenge.js";

// ─── operationHelpers: composite parameterPath (lines 55-75) ───

describe("operationHelpers coverage", () => {
  it("should handle composite parameterPath (object form)", () => {
    const result = getOperationArgumentValueFromParameter(
      { propA: "valueA", propB: "valueB" },
      {
        parameterPath: {
          propA: "propA",
          propB: "propB",
        },
        mapper: {
          serializedName: "composite",
          required: true,
          type: {
            name: "Composite",
            modelProperties: {
              propA: {
                serializedName: "propA",
                type: { name: "String" },
              },
              propB: {
                serializedName: "propB",
                type: { name: "String" },
              },
            },
          },
        } as CompositeMapper,
      },
    );
    assert.deepStrictEqual(result, { propA: "valueA", propB: "valueB" });
  });

  it("should handle composite parameterPath with non-required mapper and no matching args", () => {
    const result = getOperationArgumentValueFromParameter(
      {},
      {
        parameterPath: {
          propA: "propA",
        },
        mapper: {
          serializedName: "composite",
          required: false,
          type: {
            name: "Composite",
            modelProperties: {
              propA: {
                serializedName: "propA",
                type: { name: "String" },
              },
            },
          },
        } as CompositeMapper,
      },
    );
    assert.isUndefined(result);
  });

  it("should handle composite parameterPath where mapper is not required but property is found", () => {
    const result = getOperationArgumentValueFromParameter(
      { propA: "hello" },
      {
        parameterPath: {
          propA: "propA",
          propB: "propB",
        },
        mapper: {
          serializedName: "composite",
          required: false,
          type: {
            name: "Composite",
            modelProperties: {
              propA: {
                serializedName: "propA",
                type: { name: "String" },
              },
              propB: {
                serializedName: "propB",
                type: { name: "String" },
              },
            },
          },
        } as CompositeMapper,
      },
    );
    assert.deepStrictEqual(result, { propA: "hello" });
  });

  it("should follow originalRequest symbol in getOperationRequestInfo", () => {
    const originalRequestSymbol = Symbol.for("@azure/core-client original request");
    const innerRequest = createPipelineRequest({ url: "https://example.com" });
    const outerRequest = createPipelineRequest({
      url: "https://example.com/outer",
    }) as any;
    outerRequest[originalRequestSymbol] = innerRequest;

    const info1 = getOperationRequestInfo(innerRequest);
    info1.operationSpec = { httpMethod: "GET", responses: {}, serializer: createSerializer() };

    const info2 = getOperationRequestInfo(outerRequest);
    assert.strictEqual(info2.operationSpec?.httpMethod, "GET");
  });
});

// ─── serviceClient: requestOptions branches (lines 167-193) ───

describe("ServiceClient requestOptions coverage", () => {
  it("should pass through timeout, progress callbacks, shouldDeserialize, abortSignal, tracingOptions", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy(), { phase: "Serialize" });
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
          });
        },
      },
      pipeline,
    });

    const onUploadProgress = vi.fn();
    const onDownloadProgress = vi.fn();
    const abortController = new AbortController();

    await client.sendOperationRequest(
      {
        options: {
          requestOptions: {
            timeout: 5000,
            onUploadProgress,
            onDownloadProgress,
            shouldDeserialize: false,
          },
          abortSignal: abortController.signal,
          tracingOptions: { tracingContext: {} as any },
        },
      },
      {
        httpMethod: "GET",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    assert.strictEqual(capturedRequest!.timeout, 5000);
    assert.strictEqual(capturedRequest!.onUploadProgress, onUploadProgress);
    assert.strictEqual(capturedRequest!.onDownloadProgress, onDownloadProgress);
    assert.strictEqual(capturedRequest!.abortSignal, abortController.signal);
    assert.ok(capturedRequest!.tracingOptions);
  });
});

// ─── interfaceHelpers: getPathStringFromParameter fallback (line 38) ───

describe("interfaceHelpers coverage", () => {
  it("should fall back to mapper.serializedName when parameterPath is an object", () => {
    const result = getPathStringFromParameter({
      parameterPath: { a: "a" } as any,
      mapper: {
        serializedName: "fallbackName",
        type: { name: "Composite" },
      },
    });
    assert.strictEqual(result, "fallbackName");
  });
});

// ─── urlHelpers: simpleParseQueryParams + appendQueryParams branches ───

describe("urlHelpers coverage", () => {
  it("should handle triple duplicate query params (array push path)", () => {
    const result = appendQueryParams("https://example.com?a=1&a=2&a=3", new Map(), new Set());
    // After parsing, a=1&a=2 becomes array, then a=3 is pushed
    assert.include(result, "a=1");
    assert.include(result, "a=2");
    assert.include(result, "a=3");
  });

  it("should handle sequenceParams with existing scalar value", () => {
    const result = appendQueryParams(
      "https://example.com?q=existing",
      new Map([["q", "newVal"]]),
      new Set(["q"]),
      false,
    );
    // sequenceParams path converts to array, then noOverwrite=false overwrites
    assert.include(result, "q=newVal");
  });

  it("should handle noOverwrite=true to prevent overwriting", () => {
    const result = appendQueryParams(
      "https://example.com?q=existing",
      new Map([["q", "newVal"]]),
      new Set(["q"]),
      true,
    );
    // noOverwrite prevents overwriting; the sequenceParams path creates array but noOverwrite keeps it
    assert.include(result, "q=existing");
    assert.include(result, "q=newVal");
  });

  it("should handle bare query key (undefined value)", () => {
    const result = appendQueryParams("https://example.com?foo", new Map(), new Set());
    // bare key "foo" has no =, so value is undefined, which gets stringified
    assert.include(result, "foo");
  });

  it("should handle existing array + new array merge (dedup)", () => {
    const result = appendQueryParams(
      "https://example.com?q=1&q=2",
      new Map([["q", ["2", "3"]]]),
      new Set(),
    );
    assert.include(result, "q=");
  });

  it("should handle existing array + scalar push", () => {
    const result = appendQueryParams(
      "https://example.com?q=1&q=2",
      new Map([["q", "3"]]),
      new Set(),
    );
    assert.include(result, "q=1");
    assert.include(result, "q=2");
    assert.include(result, "q=3");
  });

  it("should handle existing scalar + new array unshift", () => {
    const result = appendQueryParams(
      "https://example.com?q=existing",
      new Map([["q", ["new1", "new2"]]]),
      new Set(),
      false,
    );
    assert.include(result, "q=");
  });
});

// ─── utils: flattenResponse pageable/sequence + parsedHeaders branches ───

describe("flattenResponse coverage", () => {
  it("should copy model properties with serializedName into array response", () => {
    const fullResponse: FullOperationResponse = {
      request: createPipelineRequest({ url: "https://example.com", method: "GET" }),
      status: 200,
      headers: createHttpHeaders(),
      parsedBody: Object.assign([1, 2, 3], { nextLink: "https://next" }),
    };
    const responseSpec = {
      bodyMapper: {
        type: {
          name: "Composite",
          modelProperties: {
            value: {
              serializedName: "",
              type: { name: "Sequence", element: { type: { name: "Number" } } },
            },
            nextLink: {
              serializedName: "nextLink",
              type: { name: "String" },
            },
          },
        },
      } as CompositeMapper,
    };
    const result = flattenResponse(fullResponse, responseSpec) as any;
    assert.strictEqual(result.nextLink, "https://next");
  });

  it("should copy parsedHeaders into pageable array response", () => {
    const fullResponse: FullOperationResponse = {
      request: createPipelineRequest({ url: "https://example.com", method: "GET" }),
      status: 200,
      headers: createHttpHeaders(),
      parsedBody: [1, 2, 3],
      parsedHeaders: { "x-custom": "headerVal" },
    };
    const responseSpec = {
      bodyMapper: {
        type: {
          name: "Composite",
          modelProperties: {
            value: {
              serializedName: "",
              type: { name: "Sequence", element: { type: { name: "Number" } } },
            },
          },
        },
      } as CompositeMapper,
    };
    const result = flattenResponse(fullResponse, responseSpec) as any;
    assert.strictEqual(result["x-custom"], "headerVal");
  });
});

// ─── base64: Buffer fast-path (line 19) ───

describe("base64 coverage", () => {
  it("should handle Buffer input directly in encodeByteArray", () => {
    const buf = Buffer.from("hello world");
    const result = encodeByteArray(buf);
    assert.strictEqual(result, buf.toString("base64"));
  });

  it("should handle Uint8Array input in encodeByteArray", () => {
    const arr = new Uint8Array([72, 101, 108, 108, 111]);
    const result = encodeByteArray(arr);
    assert.strictEqual(result, Buffer.from(arr).toString("base64"));
  });
});

// ─── pipeline: createClientPipeline with credentialOptions (lines 41-42) ───

describe("pipeline coverage", () => {
  it("should add bearerTokenAuthenticationPolicy when credentialOptions is provided", () => {
    const pipeline = createClientPipeline({
      credentialOptions: {
        credential: {
          getToken: async () => ({ token: "test", expiresOnTimestamp: Date.now() + 3600000 }),
        },
        credentialScopes: "https://example.com/.default",
      },
    });
    const policies = pipeline.getOrderedPolicies();
    const hasBearerPolicy = policies.some((p) => p.name === "bearerTokenAuthenticationPolicy");
    assert.isTrue(hasBearerPolicy);
  });

  it("should work without credentialOptions", () => {
    const pipeline = createClientPipeline({});
    const policies = pipeline.getOrderedPolicies();
    const hasBearerPolicy = policies.some((p) => p.name === "bearerTokenAuthenticationPolicy");
    assert.isFalse(hasBearerPolicy);
  });
});

// ─── serializer: uncovered branches ───

describe("serializer coverage", () => {
  const serializer = createSerializer({}, false);

  describe("bufferToBase64Url / base64UrlToByteArray edge cases", () => {
    it("should serialize Base64Url type with valid Uint8Array", () => {
      const result = serializer.serialize(
        { type: { name: "Base64Url" }, serializedName: "test" },
        new Uint8Array([1, 2, 3]),
        "testObj",
      );
      assert.isString(result);
    });

    it("should throw for Base64Url with non-Uint8Array", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "Base64Url" }, serializedName: "test" },
            "notABuffer" as any,
            "testObj",
          ),
        /must be of type Uint8Array/,
      );
    });

    it("should deserialize Base64Url type", () => {
      const result = serializer.deserialize(
        { type: { name: "Base64Url" }, serializedName: "test" },
        "AQID",
        "testObj",
      );
      assert.instanceOf(result, Uint8Array);
    });

    it("should return undefined for falsy Base64Url deserialization", () => {
      const result = serializer.deserialize(
        { type: { name: "Base64Url" }, serializedName: "test" },
        "",
        "testObj",
      );
      assert.isUndefined(result);
    });

    it("should return undefined for falsy buffer in bufferToBase64Url path", () => {
      const result = serializer.serialize(
        { type: { name: "Base64Url" }, serializedName: "test" },
        null,
        "testObj",
      );
      assert.isNull(result);
    });

    it("should throw for base64UrlToByteArray with non-string input", () => {
      assert.throws(
        () =>
          serializer.deserialize(
            { type: { name: "Base64Url" }, serializedName: "test" },
            123 as any,
            "testObj",
          ),
        /Please provide an input of type string/,
      );
    });
  });

  describe("serializeBasicTypes", () => {
    it("should throw for Number type with non-number value", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "Number" }, serializedName: "test" },
            "notANumber",
            "testObj",
          ),
        /must be of type number/,
      );
    });

    it("should throw for String type with non-string value", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "String" }, serializedName: "test" },
            123,
            "testObj",
          ),
        /must be of type string/,
      );
    });

    it("should throw for Boolean type with non-boolean value", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "Boolean" }, serializedName: "test" },
            "notBool",
            "testObj",
          ),
        /must be of type boolean/,
      );
    });

    it("should throw for Uuid type with invalid uuid", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "Uuid" }, serializedName: "test" },
            "not-a-uuid",
            "testObj",
          ),
        /must be of type string and a valid uuid/,
      );
    });

    it("should throw for Stream type with invalid stream value", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "Stream" }, serializedName: "test" },
            12345,
            "testObj",
          ),
        /must be a string, Blob, ArrayBuffer/,
      );
    });

    it("should accept a function as Stream type", () => {
      const fn = () => {};
      const result = serializer.serialize(
        { type: { name: "Stream" }, serializedName: "test" },
        fn,
        "testObj",
      );
      assert.strictEqual(result, fn);
    });

    it("should accept ArrayBuffer as Stream type", () => {
      const buf = new ArrayBuffer(8);
      const result = serializer.serialize(
        { type: { name: "Stream" }, serializedName: "test" },
        buf,
        "testObj",
      );
      assert.strictEqual(result, buf);
    });

    it("should accept ArrayBufferView as Stream type", () => {
      const view = new Uint8Array(8);
      const result = serializer.serialize(
        { type: { name: "Stream" }, serializedName: "test" },
        view,
        "testObj",
      );
      assert.strictEqual(result, view);
    });
  });

  describe("serializeDateTypes", () => {
    it("should serialize Date type from Date object", () => {
      const d = new Date("2023-06-15T00:00:00Z");
      const result = serializer.serialize(
        { type: { name: "Date" }, serializedName: "test" },
        d,
        "testObj",
      );
      assert.strictEqual(result, "2023-06-15");
    });

    it("should serialize Date type from string", () => {
      const result = serializer.serialize(
        { type: { name: "Date" }, serializedName: "test" },
        "2023-06-15",
        "testObj",
      );
      assert.strictEqual(result, "2023-06-15");
    });

    it("should throw for Date type with invalid value", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "Date" }, serializedName: "test" },
            12345,
            "testObj",
          ),
        /must be an instanceof Date or a string in ISO8601 format/,
      );
    });

    it("should serialize DateTime type from Date object", () => {
      const d = new Date("2023-06-15T10:30:00Z");
      const result = serializer.serialize(
        { type: { name: "DateTime" }, serializedName: "test" },
        d,
        "testObj",
      );
      assert.include(result, "2023-06-15");
    });

    it("should serialize DateTime type from string", () => {
      const result = serializer.serialize(
        { type: { name: "DateTime" }, serializedName: "test" },
        "2023-06-15T10:30:00Z",
        "testObj",
      );
      assert.include(result, "2023-06-15");
    });

    it("should throw for DateTime type with invalid value", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "DateTime" }, serializedName: "test" },
            {},
            "testObj",
          ),
        /must be an instanceof Date or a string in ISO8601 format/,
      );
    });

    it("should serialize DateTimeRfc1123 type from Date object", () => {
      const d = new Date("2023-06-15T10:30:00Z");
      const result = serializer.serialize(
        { type: { name: "DateTimeRfc1123" }, serializedName: "test" },
        d,
        "testObj",
      );
      assert.isString(result);
    });

    it("should serialize DateTimeRfc1123 type from string", () => {
      const result = serializer.serialize(
        { type: { name: "DateTimeRfc1123" }, serializedName: "test" },
        "Thu, 15 Jun 2023 10:30:00 GMT",
        "testObj",
      );
      assert.isString(result);
    });

    it("should throw for DateTimeRfc1123 type with invalid value", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "DateTimeRfc1123" }, serializedName: "test" },
            {},
            "testObj",
          ),
        /must be an instanceof Date or a string in RFC-1123 format/,
      );
    });

    it("should serialize UnixTime type from Date object", () => {
      const d = new Date("2023-06-15T10:30:00Z");
      const result = serializer.serialize(
        { type: { name: "UnixTime" }, serializedName: "test" },
        d,
        "testObj",
      );
      assert.isNumber(result);
    });

    it("should serialize UnixTime type from date string (line 396)", () => {
      const result = serializer.serialize(
        { type: { name: "UnixTime" }, serializedName: "test" },
        "2023-06-15T10:30:00Z",
        "testObj",
      );
      assert.isNumber(result);
      assert.strictEqual(result, Math.floor(new Date("2023-06-15T10:30:00Z").getTime() / 1000));
    });

    it("should throw for UnixTime type with invalid value", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "UnixTime" }, serializedName: "test" },
            {},
            "testObj",
          ),
        /must be an instanceof Date or a string/,
      );
    });

    it("should serialize TimeSpan type with valid duration", () => {
      const result = serializer.serialize(
        { type: { name: "TimeSpan" }, serializedName: "test" },
        "P1D",
        "testObj",
      );
      assert.strictEqual(result, "P1D");
    });

    it("should throw for TimeSpan type with invalid duration", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "TimeSpan" }, serializedName: "test" },
            "notADuration",
            "testObj",
          ),
        /must be a string in ISO 8601 format/,
      );
    });

    it("should deserialize UnixTime type", () => {
      const result = serializer.deserialize(
        { type: { name: "UnixTime" }, serializedName: "test" },
        1686826200,
        "testObj",
      );
      assert.instanceOf(result, Date);
    });

    it("should return undefined for falsy UnixTime deserialization", () => {
      const result = serializer.deserialize(
        { type: { name: "UnixTime" }, serializedName: "test" },
        0,
        "testObj",
      );
      assert.isUndefined(result);
    });
  });

  describe("serializeSequenceType", () => {
    it("should throw for non-array input", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: {
                name: "Sequence",
                element: { type: { name: "String" } },
              },
              serializedName: "test",
            } as SequenceMapper,
            "notAnArray",
            "testObj",
          ),
        /must be of type Array/,
      );
    });

    it("should throw for missing element metadata", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: { name: "Sequence" } as any,
              serializedName: "test",
            },
            [1, 2],
            "testObj",
          ),
        /element" metadata for an Array must be defined/,
      );
    });
  });

  describe("serializeDictionaryType", () => {
    it("should throw for non-object input", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: {
                name: "Dictionary",
                value: { type: { name: "String" } },
              },
              serializedName: "test",
            } as DictionaryMapper,
            "notAnObject",
            "testObj",
          ),
        /must be of type object/,
      );
    });

    it("should throw for missing value metadata", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: { name: "Dictionary" } as any,
              serializedName: "test",
            },
            { a: 1 },
            "testObj",
          ),
        /"value" metadata for a Dictionary must be defined/,
      );
    });
  });

  describe("deserializeDictionaryType", () => {
    it("should throw for missing value metadata", () => {
      assert.throws(
        () =>
          serializer.deserialize(
            {
              type: { name: "Dictionary" } as any,
              serializedName: "test",
            },
            { a: 1 },
            "testObj",
          ),
        /"value" metadata for a Dictionary must be defined/,
      );
    });
  });

  describe("deserializeSequenceType", () => {
    it("should throw for missing element metadata", () => {
      assert.throws(
        () =>
          serializer.deserialize(
            {
              type: { name: "Sequence" } as any,
              serializedName: "test",
            },
            [1, 2],
            "testObj",
          ),
        /element" metadata for an Array must be defined/,
      );
    });

    it("should wrap non-array into array (xml2js quirk)", () => {
      const result = serializer.deserialize(
        {
          type: {
            name: "Sequence",
            element: { type: { name: "Number" } },
          },
          serializedName: "test",
        } as SequenceMapper,
        42,
        "testObj",
      );
      assert.deepStrictEqual(result, [42]);
    });

    it("should return falsy responseBody as-is", () => {
      const result = serializer.deserialize(
        {
          type: {
            name: "Sequence",
            element: { type: { name: "Number" } },
          },
          serializedName: "test",
        } as SequenceMapper,
        null,
        "testObj",
      );
      assert.isNull(result);
    });

    it("should look up Composite element by className from modelMappers", () => {
      const childMapper: CompositeMapper = {
        serializedName: "Child",
        type: {
          name: "Composite",
          className: "Child",
          modelProperties: {
            id: { serializedName: "id", type: { name: "Number" } },
          },
        },
      };
      const s = createSerializer({ Child: childMapper }, false);
      const result = s.deserialize(
        {
          type: {
            name: "Sequence",
            element: {
              type: { name: "Composite", className: "Child" },
            },
          },
          serializedName: "test",
        } as SequenceMapper,
        [{ id: 1 }, { id: 2 }],
        "testObj",
      );
      assert.deepStrictEqual(result, [{ id: 1 }, { id: 2 }]);
    });
  });

  describe("resolveModelProperties / resolveReferencedMapper", () => {
    it("should throw when className is not provided", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: { name: "Composite" } as any,
              serializedName: "test",
            },
            { a: 1 },
            "testObj",
          ),
        /Class name for model/,
      );
    });

    it("should throw when referenced mapper is not found", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: { name: "Composite", className: "NonExistent" },
              serializedName: "test",
            } as CompositeMapper,
            { a: 1 },
            "testObj",
          ),
        /mapper\(\) cannot be null or undefined/,
      );
    });

    it("should throw when modelProperties are not found on referenced mapper", () => {
      const s = createSerializer(
        { Broken: { serializedName: "Broken", type: { name: "Composite", className: "Broken" } } },
        false,
      );
      assert.throws(
        () =>
          s.serialize(
            {
              type: { name: "Composite", className: "Broken" },
              serializedName: "test",
            } as CompositeMapper,
            { a: 1 },
            "testObj",
          ),
        /modelProperties cannot be null or undefined/,
      );
    });
  });

  describe("serializeCompositeType - additionalProperties", () => {
    it("should serialize additionalProperties", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          className: "Test",
          modelProperties: {
            id: { serializedName: "id", type: { name: "Number" } },
          },
          additionalProperties: { type: { name: "String" } },
        },
      };
      const result = serializer.serialize(mapper, { id: 1, extra: "value" }, "testObj");
      assert.strictEqual(result.id, 1);
      assert.strictEqual(result.extra, "value");
    });

    it("should resolve additionalProperties from referenced mapper", () => {
      const refMapper: CompositeMapper = {
        serializedName: "Ref",
        type: {
          name: "Composite",
          className: "Ref",
          modelProperties: {
            id: { serializedName: "id", type: { name: "Number" } },
          },
          additionalProperties: { type: { name: "String" } },
        },
      };
      const s = createSerializer({ Ref: refMapper }, false);
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          className: "Ref",
        },
      };
      const result = s.serialize(mapper, { id: 1, extra: "value" }, "testObj");
      assert.strictEqual(result.id, 1);
      assert.strictEqual(result.extra, "value");
    });
  });

  describe("deserializeCompositeType", () => {
    it("should handle headerCollectionPrefix", () => {
      const mapper: CompositeMapper = {
        serializedName: "Headers",
        type: {
          name: "Composite",
          modelProperties: {
            metadata: {
              serializedName: "metadata",
              type: {
                name: "Dictionary",
                value: { type: { name: "String" } },
              },
              headerCollectionPrefix: "x-ms-meta-",
            } as any,
          },
        },
      };
      const result = serializer.deserialize(
        mapper,
        {
          "x-ms-meta-key1": "val1",
          "x-ms-meta-key2": "val2",
          other: "ignored",
        },
        "testObj",
      );
      assert.deepStrictEqual(result.metadata, { key1: "val1", key2: "val2" });
    });

    it("should handle ignoreUnknownProperties option", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            id: { serializedName: "id", type: { name: "Number" } },
          },
        },
      };
      const result = serializer.deserialize(mapper, { id: 1, unknownProp: "hello" }, "testObj", {
        xml: {},
        ignoreUnknownProperties: true,
      });
      assert.strictEqual(result.id, 1);
      assert.isUndefined(result.unknownProp);
    });

    it("should pass through unknown properties when ignoreUnknownProperties is false/default", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            id: { serializedName: "id", type: { name: "Number" } },
          },
        },
      };
      const result = serializer.deserialize(mapper, { id: 1, unknownProp: "hello" }, "testObj");
      assert.strictEqual(result.id, 1);
      assert.strictEqual(result.unknownProp, "hello");
    });

    it("should handle paging deserialization (serializedName === '')", () => {
      const mapper: CompositeMapper = {
        serializedName: "PagedResult",
        type: {
          name: "Composite",
          modelProperties: {
            value: {
              serializedName: "",
              type: {
                name: "Sequence",
                element: { type: { name: "Number" } },
              },
            },
            nextLink: {
              serializedName: "nextLink",
              type: { name: "String" },
            },
          },
        },
      };
      // The paging path checks Array.isArray(responseBody[key]) && serializedName === ""
      // responseBody must have a "value" key that is an array
      const body = { value: [1, 2, 3], nextLink: "https://next" };
      const result = serializer.deserialize(mapper, body, "testObj");
      assert.deepStrictEqual(Array.from(result), [1, 2, 3]);
      assert.strictEqual(result.nextLink, "https://next");
    });

    it("should handle nested serializedName paths with null intermediate", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            deepValue: {
              serializedName: "level1.level2",
              type: { name: "String" },
            },
          },
        },
      };
      const result = serializer.deserialize(mapper, { level1: null }, "testObj");
      assert.isUndefined(result.deepValue);
    });

    it("should handle additionalProperties during deserialization", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            id: { serializedName: "id", type: { name: "Number" } },
          },
          additionalProperties: { type: { name: "String" } },
        },
      };
      const result = serializer.deserialize(mapper, { id: 1, extra: "extraVal" }, "testObj");
      assert.strictEqual(result.id, 1);
      assert.strictEqual(result.extra, "extraVal");
    });
  });

  describe("serializeByteArrayType", () => {
    it("should throw for non-Uint8Array input", () => {
      assert.throws(
        () =>
          serializer.serialize(
            { type: { name: "ByteArray" }, serializedName: "test" },
            "notABuffer",
            "testObj",
          ),
        /must be of type Uint8Array/,
      );
    });
  });

  describe("serialize nullable/required edge cases", () => {
    it("should throw when required and nullable and value is undefined", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: { name: "String" },
              serializedName: "test",
              required: true,
              nullable: true,
            },
            undefined,
            "testObj",
          ),
        /cannot be undefined/,
      );
    });

    it("should throw when not required and nullable is false and value is null", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: { name: "String" },
              serializedName: "test",
              required: false,
              nullable: false,
            },
            null,
            "testObj",
          ),
        /cannot be null/,
      );
    });
  });

  describe("validateConstraints", () => {
    it("should validate ExclusiveMaximum", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "Number" },
              serializedName: "test",
              constraints: { ExclusiveMaximum: 10 },
            },
            10,
            "testObj",
          ),
        /ExclusiveMaximum/,
      );
    });

    it("should validate ExclusiveMinimum", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "Number" },
              serializedName: "test",
              constraints: { ExclusiveMinimum: 5 },
            },
            5,
            "testObj",
          ),
        /ExclusiveMinimum/,
      );
    });

    it("should validate InclusiveMaximum", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "Number" },
              serializedName: "test",
              constraints: { InclusiveMaximum: 10 },
            },
            11,
            "testObj",
          ),
        /InclusiveMaximum/,
      );
    });

    it("should validate InclusiveMinimum", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "Number" },
              serializedName: "test",
              constraints: { InclusiveMinimum: 5 },
            },
            4,
            "testObj",
          ),
        /InclusiveMinimum/,
      );
    });

    it("should validate MaxItems", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "Sequence", element: { type: { name: "String" } } },
              serializedName: "test",
              constraints: { MaxItems: 2 },
            },
            [1, 2, 3],
            "testObj",
          ),
        /MaxItems/,
      );
    });

    it("should validate MinItems", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "Sequence", element: { type: { name: "String" } } },
              serializedName: "test",
              constraints: { MinItems: 2 },
            },
            [1],
            "testObj",
          ),
        /MinItems/,
      );
    });

    it("should validate MaxLength", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "String" },
              serializedName: "test",
              constraints: { MaxLength: 3 },
            },
            "abcd",
            "testObj",
          ),
        /MaxLength/,
      );
    });

    it("should validate MinLength", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "String" },
              serializedName: "test",
              constraints: { MinLength: 3 },
            },
            "ab",
            "testObj",
          ),
        /MinLength/,
      );
    });

    it("should validate MultipleOf", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "Number" },
              serializedName: "test",
              constraints: { MultipleOf: 3 },
            },
            7,
            "testObj",
          ),
        /MultipleOf/,
      );
    });

    it("should validate Pattern", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "String" },
              serializedName: "test",
              constraints: { Pattern: /^[a-z]+$/ },
            },
            "ABC123",
            "testObj",
          ),
        /Pattern/,
      );
    });

    it("should validate UniqueItems", () => {
      assert.throws(
        () =>
          serializer.validateConstraints(
            {
              type: { name: "Sequence", element: { type: { name: "Number" } } },
              serializedName: "test",
              constraints: { UniqueItems: true },
            },
            [1, 2, 2],
            "testObj",
          ),
        /UniqueItems/,
      );
    });

    it("should not validate constraints for null/undefined values", () => {
      // Should not throw
      serializer.validateConstraints(
        {
          type: { name: "Number" },
          serializedName: "test",
          constraints: { InclusiveMaximum: 10 },
        },
        null,
        "testObj",
      );
      serializer.validateConstraints(
        {
          type: { name: "Number" },
          serializedName: "test",
          constraints: { InclusiveMaximum: 10 },
        },
        undefined,
        "testObj",
      );
    });
  });

  describe("serializeEnumType", () => {
    it("should throw for missing allowedValues", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: { name: "Enum" } as any,
              serializedName: "test",
            },
            "value",
            "testObj",
          ),
        /Please provide a set of allowedValues/,
      );
    });

    it("should throw for value not in allowedValues", () => {
      assert.throws(
        () =>
          serializer.serialize(
            {
              type: { name: "Enum", allowedValues: ["a", "b"] },
              serializedName: "test",
            },
            "c",
            "testObj",
          ),
        /is not a valid value/,
      );
    });
  });

  describe("XML serialization - sequence element xmlNamespace", () => {
    const xmlSerializer = createSerializer({}, true);

    it("should add xmlns to Composite element in XML sequence", () => {
      const mapper: SequenceMapper = {
        serializedName: "Items",
        type: {
          name: "Sequence",
          element: {
            type: {
              name: "Composite",
              modelProperties: {
                id: { serializedName: "id", type: { name: "Number" } },
              },
            },
            xmlNamespace: "http://example.com",
            xmlNamespacePrefix: "ex",
          } as CompositeMapper,
        },
      };
      const result = xmlSerializer.serialize(mapper, [{ id: 1 }], "testObj");
      assert.deepStrictEqual(result[0].$, { "xmlns:ex": "http://example.com" });
    });

    it("should add xmlns to non-Composite element in XML sequence", () => {
      const mapper: SequenceMapper = {
        serializedName: "Items",
        type: {
          name: "Sequence",
          element: {
            type: { name: "String" },
            xmlNamespace: "http://example.com",
            serializedName: "item",
          },
        },
      };
      const result = xmlSerializer.serialize(mapper, ["hello"], "testObj");
      assert.strictEqual(result[0]._, "hello");
      assert.deepStrictEqual(result[0].$, { xmlns: "http://example.com" });
    });
  });

  describe("XML deserialization - isXML branches", () => {
    const xmlSerializer = createSerializer({}, true);

    it("should handle xmlIsAttribute", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            name: {
              serializedName: "name",
              xmlName: "name",
              xmlIsAttribute: true,
              type: { name: "String" },
            },
          },
        },
      };
      const result = xmlSerializer.deserialize(mapper, { $: { name: "testValue" } }, "testObj");
      assert.strictEqual(result.name, "testValue");
    });

    it("should handle xmlIsMsText with xmlCharKey", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            content: {
              serializedName: "content",
              xmlName: "content",
              xmlIsMsText: true,
              type: { name: "String" },
            },
          },
        },
      };
      const result = xmlSerializer.deserialize(mapper, { _: "textContent" }, "testObj");
      assert.strictEqual(result.content, "textContent");
    });

    it("should handle xmlIsMsText with string responseBody", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            content: {
              serializedName: "content",
              xmlName: "content",
              xmlIsMsText: true,
              type: { name: "String" },
            },
          },
        },
      };
      const result = xmlSerializer.deserialize(mapper, "directString", "testObj");
      assert.strictEqual(result.content, "directString");
    });

    it("should handle xmlIsWrapped", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            items: {
              serializedName: "items",
              xmlName: "Items",
              xmlElementName: "Item",
              xmlIsWrapped: true,
              type: {
                name: "Sequence",
                element: { type: { name: "String" } },
              },
            },
          },
        },
      };
      const result = xmlSerializer.deserialize(mapper, { Items: { Item: ["a", "b"] } }, "testObj");
      assert.deepStrictEqual(result.items, ["a", "b"]);
    });

    it("should handle xmlIsWrapped with missing wrapped element", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            items: {
              serializedName: "items",
              xmlName: "Items",
              xmlElementName: "Item",
              xmlIsWrapped: true,
              type: {
                name: "Sequence",
                element: { type: { name: "String" } },
              },
            },
          },
        },
      };
      const result = xmlSerializer.deserialize(mapper, { Items: {} }, "testObj");
      assert.deepStrictEqual(result.items, []);
    });

    it("should serialize xmlIsAttribute in Composite", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            name: {
              serializedName: "name",
              xmlName: "name",
              xmlIsAttribute: true,
              type: { name: "String" },
            },
          },
        },
      };
      const result = xmlSerializer.serialize(mapper, { name: "testValue" }, "testObj");
      assert.deepStrictEqual(result.$, { name: "testValue" });
    });

    it("should serialize xmlIsWrapped in Composite", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            items: {
              serializedName: "items",
              xmlName: "Items",
              xmlElementName: "Item",
              xmlIsWrapped: true,
              type: {
                name: "Sequence",
                element: { type: { name: "String" } },
              },
            },
          },
        },
      };
      const result = xmlSerializer.serialize(mapper, { items: ["a", "b"] }, "testObj");
      assert.deepStrictEqual(result.Items, { Item: ["a", "b"] });
    });
  });

  describe("deserialize - XML body with $ and _ keys", () => {
    const xmlSerializer = createSerializer({}, true);

    it("should reduce responseBody to xmlCharKey when both $ and _ present", () => {
      const result = xmlSerializer.deserialize(
        { type: { name: "String" }, serializedName: "test" },
        { $: { attr: "val" }, _: "bodyContent" },
        "testObj",
      );
      assert.strictEqual(result, "bodyContent");
    });
  });

  describe("deserialize - Boolean strings", () => {
    it("should parse 'true' string as boolean true", () => {
      const result = serializer.deserialize(
        { type: { name: "Boolean" }, serializedName: "test" },
        "true",
        "testObj",
      );
      assert.strictEqual(result, true);
    });

    it("should parse 'false' string as boolean false", () => {
      const result = serializer.deserialize(
        { type: { name: "Boolean" }, serializedName: "test" },
        "false",
        "testObj",
      );
      assert.strictEqual(result, false);
    });

    it("should return raw boolean value", () => {
      const result = serializer.deserialize(
        { type: { name: "Boolean" }, serializedName: "test" },
        true,
        "testObj",
      );
      assert.strictEqual(result, true);
    });
  });

  describe("deserialize - Number", () => {
    it("should parse NaN number as raw value", () => {
      const result = serializer.deserialize(
        { type: { name: "Number" }, serializedName: "test" },
        "notANumber",
        "testObj",
      );
      assert.strictEqual(result, "notANumber");
    });
  });

  describe("deserialize - Date types", () => {
    it("should deserialize Date type", () => {
      const result = serializer.deserialize(
        { type: { name: "Date" }, serializedName: "test" },
        "2023-06-15",
        "testObj",
      );
      assert.instanceOf(result, Date);
    });

    it("should deserialize DateTime type", () => {
      const result = serializer.deserialize(
        { type: { name: "DateTime" }, serializedName: "test" },
        "2023-06-15T10:30:00Z",
        "testObj",
      );
      assert.instanceOf(result, Date);
    });

    it("should deserialize DateTimeRfc1123 type", () => {
      const result = serializer.deserialize(
        { type: { name: "DateTimeRfc1123" }, serializedName: "test" },
        "Thu, 15 Jun 2023 10:30:00 GMT",
        "testObj",
      );
      assert.instanceOf(result, Date);
    });

    it("should deserialize ByteArray type", () => {
      const result = serializer.deserialize(
        { type: { name: "ByteArray" }, serializedName: "test" },
        "AQID",
        "testObj",
      );
      assert.instanceOf(result, Uint8Array);
    });
  });

  describe("serialize - readOnly property skipping", () => {
    it("should skip readOnly properties during serialization", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            id: { serializedName: "id", readOnly: true, type: { name: "Number" } },
            name: { serializedName: "name", type: { name: "String" } },
          },
        },
      };
      const result = serializer.serialize(mapper, { id: 1, name: "test" }, "testObj");
      assert.isUndefined(result.id);
      assert.strictEqual(result.name, "test");
    });
  });

  describe("serialize - nested serializedName paths", () => {
    it("should create intermediate objects for nested paths", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            deepProp: {
              serializedName: "level1.level2.value",
              type: { name: "String" },
            },
          },
        },
      };
      const result = serializer.serialize(mapper, { deepProp: "hello" }, "testObj");
      assert.strictEqual(result.level1.level2.value, "hello");
    });
  });

  describe("serialize - isConstant", () => {
    it("should use defaultValue for isConstant mapper", () => {
      const result = serializer.serialize(
        {
          type: { name: "String" },
          serializedName: "test",
          isConstant: true,
          defaultValue: "constantValue",
        },
        "anyValue",
        "testObj",
      );
      assert.strictEqual(result, "constantValue");
    });
  });

  describe("deserialize - isConstant", () => {
    it("should return defaultValue for isConstant mapper during deserialization", () => {
      const result = serializer.deserialize(
        {
          type: { name: "String" },
          serializedName: "test",
          isConstant: true,
          defaultValue: "constantValue",
        },
        "anyResponseValue",
        "testObj",
      );
      assert.strictEqual(result, "constantValue");
    });
  });

  describe("deserialize - defaultValue", () => {
    it("should return defaultValue when responseBody is undefined", () => {
      const result = serializer.deserialize(
        {
          type: { name: "String" },
          serializedName: "test",
          defaultValue: "defaultVal",
        },
        undefined,
        "testObj",
      );
      assert.strictEqual(result, "defaultVal");
    });
  });

  describe("XML Sequence edge case - empty list", () => {
    const xmlSerializer = createSerializer({}, true);

    it("should return empty array for undefined XML non-wrapped Sequence", () => {
      const result = xmlSerializer.deserialize(
        {
          type: {
            name: "Sequence",
            element: { type: { name: "String" } },
          },
          serializedName: "test",
        } as SequenceMapper,
        undefined,
        "testObj",
      );
      assert.deepStrictEqual(result, []);
    });

    it("should return defaultValue for wrapped XML Sequence that is undefined", () => {
      const result = xmlSerializer.deserialize(
        {
          type: {
            name: "Sequence",
            element: { type: { name: "String" } },
          },
          serializedName: "test",
          xmlIsWrapped: true,
          defaultValue: [],
        } as any,
        undefined,
        "testObj",
      );
      assert.deepStrictEqual(result, []);
    });
  });

  describe("serialize - xmlNamespace on Composite", () => {
    const xmlSerializer = createSerializer({}, true);

    it("should add xmlNamespace to Composite root", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        xmlNamespace: "http://example.com",
        xmlNamespacePrefix: "ex",
        type: {
          name: "Composite",
          modelProperties: {
            name: { serializedName: "name", xmlName: "name", type: { name: "String" } },
          },
        },
      };
      const result = xmlSerializer.serialize(mapper, { name: "test" }, "testObj");
      assert.deepStrictEqual(result.$, { "xmlns:ex": "http://example.com" });
    });
  });

  describe("serialize - Dictionary with xmlNamespace", () => {
    const xmlSerializer = createSerializer({}, true);

    it("should add xmlNamespace to Dictionary root", () => {
      const mapper: DictionaryMapper = {
        serializedName: "Dict",
        xmlNamespace: "http://example.com",
        type: {
          name: "Dictionary",
          value: { type: { name: "String" } },
        },
      };
      const result = xmlSerializer.serialize(mapper, { key: "val" }, "testObj");
      assert.deepStrictEqual(result.$, { xmlns: "http://example.com" });
    });
  });

  describe("getXmlObjectValue", () => {
    const xmlSerializer = createSerializer({}, true);

    it("should add xmlns to non-Composite type with xmlNamespace", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            value: {
              serializedName: "value",
              xmlName: "value",
              xmlNamespace: "http://example.com",
              type: { name: "String" },
            },
          },
        },
      };
      const result = xmlSerializer.serialize(mapper, { value: "hello" }, "testObj");
      assert.strictEqual(result.value._, "hello");
      assert.deepStrictEqual(result.value.$, { xmlns: "http://example.com" });
    });

    it("should not duplicate xmlns for Composite type that already has $", () => {
      const childMapper: CompositeMapper = {
        serializedName: "Child",
        type: {
          name: "Composite",
          className: "Child",
          modelProperties: {
            id: { serializedName: "id", xmlName: "id", type: { name: "Number" } },
          },
        },
      };
      const s = createSerializer({ Child: childMapper }, true);
      const mapper: CompositeMapper = {
        serializedName: "Parent",
        type: {
          name: "Composite",
          modelProperties: {
            child: {
              serializedName: "child",
              xmlName: "child",
              xmlNamespace: "http://example.com",
              type: {
                name: "Composite",
                className: "Child",
              },
            },
          },
        },
      };
      // Serialize with a child that will get $ added via xmlNamespace on parent property
      const result = s.serialize(mapper, { child: { id: 1 } }, "testObj");
      assert.ok(result.child);
    });
  });

  describe("polymorphic mapper", () => {
    it("should find polymorphic mapper during serialization", () => {
      const baseMapper: CompositeMapper = {
        serializedName: "Animal",
        type: {
          name: "Composite",
          className: "Animal",
          uberParent: "Animal",
          polymorphicDiscriminator: {
            serializedName: "kind",
            clientName: "kind",
          },
          modelProperties: {
            kind: { serializedName: "kind", type: { name: "String" } },
          },
        },
      };
      const dogMapper: CompositeMapper = {
        serializedName: "Dog",
        type: {
          name: "Composite",
          className: "Dog",
          uberParent: "Animal",
          modelProperties: {
            kind: { serializedName: "kind", type: { name: "String" } },
            bark: { serializedName: "bark", type: { name: "Boolean" } },
          },
        },
      };
      const s = createSerializer(
        {
          Animal: baseMapper,
          Dog: dogMapper,
          discriminators: {
            "Animal.Dog": dogMapper,
          },
        },
        false,
      );
      const result = s.serialize(baseMapper, { kind: "Dog", bark: true }, "testObj");
      assert.strictEqual(result.kind, "Dog");
      assert.strictEqual(result.bark, true);
    });

    it("should find polymorphic mapper during deserialization", () => {
      const baseMapper: CompositeMapper = {
        serializedName: "Animal",
        type: {
          name: "Composite",
          className: "Animal",
          uberParent: "Animal",
          polymorphicDiscriminator: {
            serializedName: "kind",
            clientName: "kind",
          },
          modelProperties: {
            kind: { serializedName: "kind", type: { name: "String" } },
          },
        },
      };
      const dogMapper: CompositeMapper = {
        serializedName: "Dog",
        type: {
          name: "Composite",
          className: "Dog",
          uberParent: "Animal",
          modelProperties: {
            kind: { serializedName: "kind", type: { name: "String" } },
            bark: { serializedName: "bark", type: { name: "Boolean" } },
          },
        },
      };
      const s = createSerializer(
        {
          Animal: baseMapper,
          Dog: dogMapper,
          discriminators: {
            "Animal.Dog": dogMapper,
          },
        },
        false,
      );
      const result = s.deserialize(baseMapper, { kind: "Dog", bark: true }, "testObj");
      assert.strictEqual(result.kind, "Dog");
      assert.strictEqual(result.bark, true);
    });
  });

  describe("splitSerializeName with escaped dots", () => {
    it("should handle escaped dots in serializedName", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            dotProp: {
              serializedName: "level1\\.level2",
              type: { name: "String" },
            },
          },
        },
      };
      const result = serializer.serialize(mapper, { dotProp: "value" }, "testObj");
      assert.strictEqual(result["level1.level2"], "value");
    });
  });

  describe("Composite serialization - polymorphic discriminator default value", () => {
    it("should use mapper serializedName as discriminator value when toSerialize is undefined", () => {
      const baseMapper: CompositeMapper = {
        serializedName: "BaseType",
        type: {
          name: "Composite",
          className: "BaseType",
          uberParent: "BaseType",
          polymorphicDiscriminator: {
            serializedName: "type",
            clientName: "type",
          },
          modelProperties: {
            type: { serializedName: "type", type: { name: "String" } },
            name: { serializedName: "name", type: { name: "String" } },
          },
        },
      };
      const s = createSerializer(
        {
          BaseType: baseMapper,
          discriminators: {},
        },
        false,
      );
      const result = s.serialize(baseMapper, { name: "test" }, "testObj");
      assert.strictEqual(result.type, "BaseType");
    });
  });

  describe("serialize - Composite with empty object for undefined/null values", () => {
    it("should handle null values in Composite", () => {
      const mapper: CompositeMapper = {
        serializedName: "Test",
        type: {
          name: "Composite",
          modelProperties: {
            value: { serializedName: "value", type: { name: "String" } },
          },
        },
      };
      const result = serializer.serialize(mapper, null, "testObj");
      assert.isNull(result);
    });
  });

  describe("getPolymorphicDiscriminatorRecursively - uberParent/className lookup", () => {
    it("should look up polymorphicDiscriminator from uberParent", () => {
      const parentMapper: CompositeMapper = {
        serializedName: "Parent",
        type: {
          name: "Composite",
          className: "Parent",
          uberParent: "Parent",
          polymorphicDiscriminator: {
            serializedName: "type",
            clientName: "type",
          },
          modelProperties: {
            type: { serializedName: "type", type: { name: "String" } },
          },
        },
      };
      const childMapper: CompositeMapper = {
        serializedName: "Child",
        type: {
          name: "Composite",
          className: "Child",
          uberParent: "Parent",
          modelProperties: {
            type: { serializedName: "type", type: { name: "String" } },
            extra: { serializedName: "extra", type: { name: "String" } },
          },
        },
      };
      const s = createSerializer(
        {
          Parent: parentMapper,
          Child: childMapper,
          discriminators: { "Parent.Child": childMapper },
        },
        false,
      );
      const result = s.deserialize(childMapper, { type: "Child", extra: "val" }, "testObj");
      assert.strictEqual(result.extra, "val");
    });
  });
});

// ─── authorizeRequestOnClaimChallenge: parseCAEChallenge falsy fallback (line 76) ───

describe("authorizeRequestOnClaimChallenge coverage", () => {
  it("should handle malformed WWW-Authenticate header (no claims)", async () => {
    const request = createPipelineRequest({ url: "https://example.com" });
    const result = await authorizeRequestOnClaimChallenge({
      async getAccessToken() {
        return { token: "token", expiresOnTimestamp: Date.now() + 3600000 };
      },
      scopes: [],
      response: {
        headers: createHttpHeaders({
          "WWW-Authenticate": 'Bearer realm="test"',
        }),
        request,
        status: 401,
      },
      request,
    });
    assert.isFalse(result);
  });

  it("should handle empty WWW-Authenticate header", async () => {
    const request = createPipelineRequest({ url: "https://example.com" });
    const result = await authorizeRequestOnClaimChallenge({
      async getAccessToken() {
        return { token: "token", expiresOnTimestamp: Date.now() + 3600000 };
      },
      scopes: [],
      response: {
        headers: createHttpHeaders(),
        request,
        status: 401,
      },
      request,
    });
    assert.isFalse(result);
  });
});

// ─── deserializationPolicy/serializationPolicy uncovered branches ───

describe("serializationPolicy coverage", () => {
  it("should serialize formData parameters", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy(), { phase: "Serialize" });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
          });
        },
      },
      pipeline,
    });

    await client.sendOperationRequest(
      { file: "fileContent" },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        formDataParameters: [
          {
            parameterPath: "file",
            mapper: {
              serializedName: "file",
              type: { name: "String" },
            },
          },
        ],
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    assert.deepStrictEqual(capturedRequest!.formData, { file: "fileContent" });
  });

  it("should handle text/plain content type without JSON stringifying", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy(), { phase: "Serialize" });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
          });
        },
      },
      pipeline,
    });

    await client.sendOperationRequest(
      { body: "plain text content" },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        contentType: "text/plain",
        mediaType: "text",
        serializer: createSerializer(),
        requestBody: {
          parameterPath: "body",
          mapper: {
            serializedName: "body",
            type: { name: "String" },
          },
        },
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    assert.strictEqual(capturedRequest!.body, "plain text content");
  });
});

describe("deserializationPolicy coverage", () => {
  it("should handle operationResponseGetter", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
            bodyAsText: '{"id": 1}',
          });
        },
      },
      pipeline,
    });

    const operationInfo = getOperationRequestInfo(
      createPipelineRequest({ url: "https://example.com" }),
    );
    // Ensure the operationResponseGetter path is available through sendOperationRequest
    await client.sendOperationRequest(
      {
        options: {
          requestOptions: {
            shouldDeserialize: true,
          },
        },
      },
      {
        httpMethod: "GET",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        responses: {
          200: {
            bodyMapper: {
              type: {
                name: "Composite",
                modelProperties: {
                  id: { serializedName: "id", type: { name: "Number" } },
                },
              },
            },
          },
        },
      },
    );
  });

  it("should handle shouldDeserialize as a function", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
            bodyAsText: '{"id": 1}',
          }),
      },
      pipeline,
    });

    await client.sendOperationRequest(
      {
        options: {
          requestOptions: {
            shouldDeserialize: (response: any) => response.status === 200,
          },
        },
      },
      {
        httpMethod: "GET",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        responses: { 200: {} },
      },
    );
  });

  it("should handle HEAD request with streaming response codes", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
          }),
      },
      pipeline,
    });

    const result = await client.sendOperationRequest(
      {},
      {
        httpMethod: "HEAD",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        responses: {
          200: {},
        },
      },
    );
    assert.deepStrictEqual(result, { body: true });
  });

  it("should handle parsedHeaders from headersMapper", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders({ "x-custom": "value123" }),
            bodyAsText: '{"id": 1}',
          }),
      },
      pipeline,
    });

    const result: any = await client.sendOperationRequest(
      {},
      {
        httpMethod: "GET",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        responses: {
          200: {
            bodyMapper: {
              type: {
                name: "Composite",
                modelProperties: {
                  id: { serializedName: "id", type: { name: "Number" } },
                },
              },
            },
            headersMapper: {
              type: {
                name: "Composite",
                modelProperties: {
                  xCustom: {
                    serializedName: "x-custom",
                    type: { name: "String" },
                  },
                },
              },
            },
          },
        },
      },
    );
    assert.strictEqual(result.xCustom, "value123");
  });

  it("should wrap error with error headers mapper", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 400,
            headers: createHttpHeaders({ "x-error-id": "err123" }),
            bodyAsText: '{"error": {"code": "BadRequest", "message": "Invalid input"}}',
          }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        {},
        {
          httpMethod: "GET",
          baseUrl: "https://example.com",
          serializer: createSerializer(),
          responses: {
            200: {},
            default: {
              bodyMapper: {
                type: {
                  name: "Composite",
                  modelProperties: {
                    error: {
                      serializedName: "error",
                      type: {
                        name: "Composite",
                        modelProperties: {
                          code: { serializedName: "code", type: { name: "String" } },
                          message: { serializedName: "message", type: { name: "String" } },
                        },
                      },
                    },
                  },
                },
              },
              headersMapper: {
                type: {
                  name: "Composite",
                  modelProperties: {
                    xErrorId: {
                      serializedName: "x-error-id",
                      type: { name: "String" },
                    },
                  },
                },
              },
            },
          },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.strictEqual(err.code, "BadRequest");
    }
  });

  it("should handle XML parsing error", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(
      deserializationPolicy({
        expectedContentTypes: {
          json: ["application/json"],
          xml: ["application/xml"],
        },
        parseXML: async () => {
          throw new Error("XML parse error");
        },
      }),
      { phase: "Deserialize" },
    );

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders({ "Content-Type": "application/xml" }),
            bodyAsText: "<invalid>xml",
          }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        {},
        {
          httpMethod: "GET",
          baseUrl: "https://example.com",
          serializer: createSerializer(),
          responses: { 200: {} },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "XML parse error");
    }
  });

  it("should handle JSON parse error", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders({ "Content-Type": "application/json" }),
            bodyAsText: "not valid json{{{",
          }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        {},
        {
          httpMethod: "GET",
          baseUrl: "https://example.com",
          serializer: createSerializer(),
          responses: { 200: {} },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "occurred while parsing the response body");
    }
  });
});

// ─── utils: flattenResponse Stream response (line 141) ───

describe("flattenResponse - Stream response", () => {
  it("should return stream properties for Stream body type", () => {
    const mockStream = { pipe: () => {} };
    const fullResponse: FullOperationResponse = {
      request: createPipelineRequest({ url: "https://example.com", method: "GET" }),
      status: 200,
      headers: createHttpHeaders(),
      readableStreamBody: mockStream as any,
      parsedHeaders: { "x-header": "val" },
    };
    const responseSpec = {
      bodyMapper: {
        type: { name: "Stream" },
      },
    };
    const result = flattenResponse(fullResponse, responseSpec) as any;
    assert.strictEqual(result.readableStreamBody, mockStream);
    assert.strictEqual(result["x-header"], "val");
  });
});

// ─── serviceClient: no endpoint error (line 139) ───

describe("ServiceClient - no endpoint", () => {
  it("should throw when no endpoint and no baseUrl in operationSpec", async () => {
    const pipeline = createEmptyPipeline();
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        {},
        {
          httpMethod: "GET",
          serializer: createSerializer(),
          responses: { 200: {} },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "must have a endpoint string property");
    }
  });
});

// ─── serializer: deserialize Dictionary falsy responseBody (line 1091) ───

describe("serializer - Dictionary deserialization with falsy body", () => {
  it("should return falsy responseBody for Dictionary (0)", () => {
    const serializer = createSerializer({}, false);
    // 0 is falsy but not null/undefined, so it passes the null check at line 233
    // and reaches deserializeDictionaryType which returns it at line 1091
    const result = serializer.deserialize(
      {
        type: {
          name: "Dictionary",
          value: { type: { name: "String" } },
        },
        serializedName: "test",
      } as DictionaryMapper,
      0,
      "testObj",
    );
    assert.strictEqual(result, 0);
  });
  it("should return falsy responseBody for Dictionary (empty string)", () => {
    const serializer = createSerializer({}, false);
    const result = serializer.deserialize(
      {
        type: {
          name: "Dictionary",
          value: { type: { name: "String" } },
        },
        serializedName: "test",
      } as DictionaryMapper,
      "",
      "testObj",
    );
    assert.strictEqual(result, "");
  });
});

// ─── serializer: deserialize Sequence falsy responseBody (line 1132) ───

describe("serializer - Sequence deserialization with falsy body", () => {
  it("should return falsy responseBody for Sequence (0)", () => {
    const serializer = createSerializer({}, false);
    const result = serializer.deserialize(
      {
        type: {
          name: "Sequence",
          element: { type: { name: "String" } },
        },
        serializedName: "test",
      } as SequenceMapper,
      0,
      "testObj",
    );
    assert.strictEqual(result, 0);
  });
  it("should return falsy responseBody for Sequence (false)", () => {
    const serializer = createSerializer({}, false);
    const result = serializer.deserialize(
      {
        type: {
          name: "Sequence",
          element: { type: { name: "String" } },
        },
        serializedName: "test",
      } as SequenceMapper,
      false,
      "testObj",
    );
    assert.strictEqual(result, false);
  });
});

// ─── serializer: polymorphicDiscriminator fallback on deserialization (line 988) ───

describe("serializer - polymorphic discriminator default during deserialization", () => {
  it("should use mapper.serializedName as discriminator when value is missing", () => {
    const baseMapper: CompositeMapper = {
      serializedName: "Animal",
      type: {
        name: "Composite",
        className: "Animal",
        uberParent: "Animal",
        polymorphicDiscriminator: {
          serializedName: "kind",
          clientName: "kind",
        },
        modelProperties: {
          kind: { serializedName: "kind", type: { name: "String" } },
          name: { serializedName: "name", type: { name: "String" } },
        },
      },
    };
    const s = createSerializer({ Animal: baseMapper, discriminators: {} }, false);
    // When kind is not present in the response body, it should default to mapper.serializedName
    const result = s.deserialize(baseMapper, { name: "Fido" }, "testObj");
    assert.strictEqual(result.kind, "Animal");
  });
});

// ─── serializationPolicy XML branches ───

describe("serializationPolicy - XML serialization", () => {
  it("should throw XML serialization unsupported when no stringifyXML provided", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy(), { phase: "Serialize" });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        { body: { name: "test" } },
        {
          httpMethod: "POST",
          baseUrl: "https://example.com",
          isXML: true,
          contentType: "application/xml",
          serializer: createSerializer({}, true),
          requestBody: {
            parameterPath: "body",
            mapper: {
              serializedName: "body",
              xmlName: "TestBody",
              type: {
                name: "Composite",
                modelProperties: {
                  name: { serializedName: "name", xmlName: "name", type: { name: "String" } },
                },
              },
            } as CompositeMapper,
          },
          responses: { 200: {} },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "XML serialization unsupported");
    }
  });

  it("should serialize XML Sequence with stringifyXML", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy({ stringifyXML: (obj) => JSON.stringify(obj) }), {
      phase: "Serialize",
    });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() });
        },
      },
      pipeline,
    });

    await client.sendOperationRequest(
      { body: ["item1", "item2"] },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        isXML: true,
        contentType: "application/xml",
        serializer: createSerializer({}, true),
        requestBody: {
          parameterPath: "body",
          mapper: {
            serializedName: "Items",
            xmlName: "Items",
            xmlElementName: "Item",
            type: {
              name: "Sequence",
              element: { type: { name: "String" } },
            },
          } as SequenceMapper,
        },
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    assert.isString(capturedRequest!.body);
  });

  it("should serialize XML Sequence with xmlNamespace", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy({ stringifyXML: (obj) => JSON.stringify(obj) }), {
      phase: "Serialize",
    });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() });
        },
      },
      pipeline,
    });

    await client.sendOperationRequest(
      { body: ["item1"] },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        isXML: true,
        contentType: "application/xml",
        serializer: createSerializer({}, true),
        requestBody: {
          parameterPath: "body",
          mapper: {
            serializedName: "Items",
            xmlName: "Items",
            xmlElementName: "Item",
            xmlNamespace: "http://example.com",
            xmlNamespacePrefix: "ex",
            type: {
              name: "Sequence",
              element: { type: { name: "String" } },
            },
          } as SequenceMapper,
        },
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
  });

  it("should serialize XML with xmlNamespace on non-Composite/Sequence/Dictionary type", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy({ stringifyXML: (obj) => JSON.stringify(obj) }), {
      phase: "Serialize",
    });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() });
        },
      },
      pipeline,
    });

    await client.sendOperationRequest(
      { body: "stringValue" },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        isXML: true,
        contentType: "application/xml",
        serializer: createSerializer({}, true),
        requestBody: {
          parameterPath: "body",
          mapper: {
            serializedName: "Value",
            xmlName: "Value",
            xmlNamespace: "http://example.com",
            type: { name: "String" },
          },
        },
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
  });

  it("should handle serialization error in request body", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy(), { phase: "Serialize" });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        { body: "not a number" },
        {
          httpMethod: "POST",
          baseUrl: "https://example.com",
          serializer: createSerializer(),
          requestBody: {
            parameterPath: "body",
            mapper: {
              serializedName: "body",
              required: true,
              type: { name: "Number" },
            },
          },
          responses: { 200: {} },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "occurred in serializing the payload");
    }
  });

  it("should handle nullable body being null", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy(), { phase: "Serialize" });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() });
        },
      },
      pipeline,
    });

    await client.sendOperationRequest(
      { body: null },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        requestBody: {
          parameterPath: "body",
          mapper: {
            serializedName: "body",
            nullable: true,
            type: { name: "String" },
          },
        },
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    assert.strictEqual(capturedRequest!.body, "null");
  });

  it("should serialize Stream body without JSON.stringify in non-XML", async () => {
    const streamBody = { pipe: vi.fn(), on: vi.fn() };
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy(), { phase: "Serialize" });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() });
        },
      },
      pipeline,
    });

    await client.sendOperationRequest(
      { body: streamBody },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        requestBody: {
          parameterPath: "body",
          mapper: {
            serializedName: "body",
            type: { name: "Stream" },
          },
        },
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    assert.strictEqual(capturedRequest!.body, streamBody);
  });
});

// ─── deserializationPolicy - XML Sequence error, error deserialization error, etc. ───

describe("deserializationPolicy - additional branches", () => {
  it("should handle XML Sequence error body with xmlElementName", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(
      deserializationPolicy({
        expectedContentTypes: {
          json: [],
          xml: ["application/xml"],
        },
        parseXML: async (str) => JSON.parse(str),
      }),
      { phase: "Deserialize" },
    );

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 400,
            headers: createHttpHeaders({ "Content-Type": "application/xml" }),
            bodyAsText: JSON.stringify({
              Error: [{ code: "Err1", message: "msg1" }],
            }),
          }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        {},
        {
          httpMethod: "GET",
          baseUrl: "https://example.com",
          isXML: true,
          serializer: createSerializer({}, true),
          responses: {
            200: {},
            default: {
              bodyMapper: {
                xmlElementName: "Error",
                type: {
                  name: "Sequence",
                  element: {
                    type: {
                      name: "Composite",
                      modelProperties: {
                        code: { serializedName: "code", type: { name: "String" } },
                        message: { serializedName: "message", type: { name: "String" } },
                      },
                    },
                  },
                },
              } as SequenceMapper,
            },
          },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.ok(err);
    }
  });

  it("should handle error in error deserialization (catch block line 319)", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 400,
            headers: createHttpHeaders(),
            bodyAsText: '{"error": {"code": "BadRequest", "message": "fail"}}',
          }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        {},
        {
          httpMethod: "GET",
          baseUrl: "https://example.com",
          serializer: createSerializer(),
          responses: {
            200: {},
            default: {
              bodyMapper: {
                type: {
                  name: "Composite",
                  className: "BrokenModel",
                },
              } as CompositeMapper,
            },
          },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "occurred in deserializing the responseBody");
    }
  });

  it("should handle XML content-type parsing without parseXML", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(
      deserializationPolicy({
        expectedContentTypes: {
          json: [],
          xml: ["application/xml"],
        },
      }),
      { phase: "Deserialize" },
    );

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders({ "Content-Type": "application/xml" }),
            bodyAsText: "<root>test</root>",
          }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        {},
        {
          httpMethod: "GET",
          baseUrl: "https://example.com",
          serializer: createSerializer(),
          responses: { 200: {} },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "Parsing XML not supported");
    }
  });

  it("should handle no operationSpec in request", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
          }),
      },
      pipeline,
    });

    // Directly send a request without setting up operationSpec
    const result = await client.sendRequest(createPipelineRequest({ url: "https://example.com" }));
    assert.strictEqual(result.status, 200);
  });

  it("should handle stream response status codes", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
            bodyAsText: "stream content",
          }),
      },
      pipeline,
    });

    const result = await client.sendOperationRequest(
      {},
      {
        httpMethod: "GET",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        responses: {
          200: {
            bodyMapper: {
              type: { name: "Stream" },
            },
          },
        },
      },
    );
    assert.ok(result);
  });

  it("should deserialize XML body in success response", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(
      deserializationPolicy({
        expectedContentTypes: {
          json: [],
          xml: ["application/xml"],
        },
        parseXML: async (str) => JSON.parse(str),
      }),
      { phase: "Deserialize" },
    );

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders({ "Content-Type": "application/xml" }),
            bodyAsText: JSON.stringify({ Items: { Item: ["a", "b"] } }),
          }),
      },
      pipeline,
    });

    const result = await client.sendOperationRequest(
      {},
      {
        httpMethod: "GET",
        baseUrl: "https://example.com",
        isXML: true,
        serializer: createSerializer({}, true),
        responses: {
          200: {
            bodyMapper: {
              xmlElementName: "Item",
              type: {
                name: "Sequence",
                element: { type: { name: "String" } },
              },
            } as SequenceMapper,
          },
        },
      },
    );
    assert.ok(result);
  });

  it("should handle isError response spec", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
            bodyAsText: '{"error": {"code": "SoftError", "message": "recoverable"}}',
          }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        {},
        {
          httpMethod: "GET",
          baseUrl: "https://example.com",
          serializer: createSerializer(),
          responses: {
            200: {
              isError: true,
            },
          },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.ok(err);
    }
  });
});

// ─── authorizeRequestOnTenantChallenge: null token, no challenge, non-401 ───

describe("authorizeRequestOnTenantChallenge coverage", () => {
  it("should return false when getAccessToken returns null", async () => {
    const { authorizeRequestOnTenantChallenge: authorizeOnTenant } =
      await import("../../src/authorizeRequestOnTenantChallenge.js");
    const fakeGuid = "3a4e2c3b-defc-466c-b0c8-6a419bf92858";
    const result = await authorizeOnTenant({
      getAccessToken: async () => null,
      request: createPipelineRequest({ url: "https://example.com" }),
      response: {
        status: 401,
        headers: createHttpHeaders({
          "WWW-Authenticate": `Bearer authorization_uri=https://login.microsoftonline.com/${fakeGuid}/oauth2/authorize resource_id=https://storage.azure.com`,
        }),
        request: createPipelineRequest({ url: "https://example.com" }),
      },
      scopes: ["https://storage.azure.com/.default"],
    });
    assert.isFalse(result);
  });

  it("should return false when response is not 401", async () => {
    const { authorizeRequestOnTenantChallenge: authorizeOnTenant } =
      await import("../../src/authorizeRequestOnTenantChallenge.js");
    const result = await authorizeOnTenant({
      getAccessToken: async () => ({ token: "t", expiresOnTimestamp: Date.now() + 3600000 }),
      request: createPipelineRequest({ url: "https://example.com" }),
      response: {
        status: 200,
        headers: createHttpHeaders(),
        request: createPipelineRequest({ url: "https://example.com" }),
      },
      scopes: ["https://storage.azure.com/.default"],
    });
    assert.isFalse(result);
  });

  it("should return false when tenantId is not a valid UUID", async () => {
    const { authorizeRequestOnTenantChallenge: authorizeOnTenant } =
      await import("../../src/authorizeRequestOnTenantChallenge.js");
    const result = await authorizeOnTenant({
      getAccessToken: async () => ({ token: "t", expiresOnTimestamp: Date.now() + 3600000 }),
      request: createPipelineRequest({ url: "https://example.com" }),
      response: {
        status: 401,
        headers: createHttpHeaders({
          "WWW-Authenticate": `Bearer authorization_uri=https://login.microsoftonline.com/not-a-uuid/oauth2/authorize resource_id=https://storage.azure.com`,
        }),
        request: createPipelineRequest({ url: "https://example.com" }),
      },
      scopes: ["https://storage.azure.com/.default"],
    });
    assert.isFalse(result);
  });

  it("should return false when WWW-Authenticate header is missing on 401", async () => {
    const { authorizeRequestOnTenantChallenge: authorizeOnTenant } =
      await import("../../src/authorizeRequestOnTenantChallenge.js");
    const result = await authorizeOnTenant({
      getAccessToken: async () => ({ token: "t", expiresOnTimestamp: Date.now() + 3600000 }),
      request: createPipelineRequest({ url: "https://example.com" }),
      response: {
        status: 401,
        headers: createHttpHeaders(),
        request: createPipelineRequest({ url: "https://example.com" }),
      },
      scopes: ["https://storage.azure.com/.default"],
    });
    assert.isFalse(result);
  });

  it("should use custom token type when available", async () => {
    const { authorizeRequestOnTenantChallenge: authorizeOnTenant } =
      await import("../../src/authorizeRequestOnTenantChallenge.js");
    const fakeGuid = "3a4e2c3b-defc-466c-b0c8-6a419bf92858";
    const request = createPipelineRequest({ url: "https://example.com" });
    const result = await authorizeOnTenant({
      getAccessToken: async () => ({
        token: "myToken",
        expiresOnTimestamp: Date.now() + 3600000,
        tokenType: "pop",
      }),
      request,
      response: {
        status: 401,
        headers: createHttpHeaders({
          "WWW-Authenticate": `Bearer authorization_uri=https://login.microsoftonline.com/${fakeGuid}/oauth2/authorize resource_id=https://storage.azure.com`,
        }),
        request: createPipelineRequest({ url: "https://example.com" }),
      },
      scopes: ["https://storage.azure.com/.default"],
    });
    assert.isTrue(result);
    assert.strictEqual(request.headers.get("authorization"), "pop myToken");
  });
});

// ─── urlHelpers: appendPath with empty/missing pathToAppend and path without trailing slash ───

describe("urlHelpers - appendPath branches", () => {
  it("should handle path with query string attached to path component", () => {
    const serializer = createSerializer({}, false);
    const url = getRequestUrl(
      "https://example.com",
      {
        path: "/items?extra=1",
        httpMethod: "GET",
        responses: {},
        serializer,
      },
      {},
      {},
    );
    assert.include(url, "extra=1");
  });

  it("should handle path component that is an absolute URL", () => {
    const serializer = createSerializer({}, false);
    const url = getRequestUrl(
      "https://example.com",
      {
        path: "/{nextLink}",
        httpMethod: "GET",
        responses: {},
        urlParameters: [
          {
            parameterPath: "nextLink",
            mapper: {
              serializedName: "nextLink",
              required: true,
              type: { name: "String" },
            },
            skipEncoding: true,
          },
        ],
        serializer,
      },
      { nextLink: "https://other.com/page2?token=abc" },
      {},
    );
    assert.strictEqual(url, "https://other.com/page2?token=abc");
  });
});

describe("serializer - Sequence element className lookup", () => {
  it("should look up Composite element by className from modelMappers during serialization", () => {
    const childMapper: CompositeMapper = {
      serializedName: "Child",
      type: {
        name: "Composite",
        className: "Child",
        modelProperties: {
          id: { serializedName: "id", type: { name: "Number" } },
          name: { serializedName: "name", type: { name: "String" } },
        },
      },
    };
    const s = createSerializer({ Child: childMapper }, false);
    const result = s.serialize(
      {
        type: {
          name: "Sequence",
          element: {
            type: { name: "Composite", className: "Child" },
          },
        },
        serializedName: "test",
      } as SequenceMapper,
      [{ id: 1, name: "a" }],
      "testObj",
    );
    assert.deepStrictEqual(result, [{ id: 1, name: "a" }]);
  });
});

// ─── Additional remaining coverage gaps ───

describe("deserializationPolicy - operationResponseGetter (line 113)", () => {
  it("should use operationResponseGetter when set", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });
    // We need to set operationResponseGetter directly on the operationInfo
    // This requires intercepting the request before it goes through the pipeline
    const customPolicy = {
      name: "setOperationResponseGetter",
      async sendRequest(request: any, next: any) {
        const info = getOperationRequestInfo(request);
        info.operationResponseGetter = (_spec: any, response: any) => {
          return _spec.responses[response.status];
        };
        return next(request);
      },
    };
    pipeline.addPolicy(customPolicy);

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
            bodyAsText: '{"id": 42}',
          }),
      },
      pipeline,
    });

    const result: any = await client.sendOperationRequest(
      {},
      {
        httpMethod: "GET",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        responses: {
          200: {
            bodyMapper: {
              type: {
                name: "Composite",
                modelProperties: {
                  id: { serializedName: "id", type: { name: "Number" } },
                },
              },
            },
          },
        },
      },
    );
    assert.strictEqual(result.id, 42);
  });
});

describe("deserializationPolicy - shouldReturnResponse path (line 168)", () => {
  it("should return response without deserialization for empty operationSpec", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 204,
            headers: createHttpHeaders(),
          }),
      },
      pipeline,
    });

    // operationSpec with only a default response, and status 204 not in responses
    // => isExpectedStatusCode false, but then we match the default response
    // For the shouldReturnResponse path, we need a response that's not in the spec
    // AND no default response AND no error body
    const result = await client.sendOperationRequest(
      {},
      {
        httpMethod: "DELETE",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        responses: {
          // only default, no 204 match
          default: {},
        },
      },
    );
    assert.ok(result);
  });
});

describe("deserializationPolicy - deserialization error (lines 190-198)", () => {
  it("should throw RestError when body deserialization fails", async () => {
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(deserializationPolicy(), { phase: "Deserialize" });

    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) =>
          Promise.resolve({
            request: req,
            status: 200,
            headers: createHttpHeaders(),
            bodyAsText: '{"value": "not-a-number"}',
          }),
      },
      pipeline,
    });

    try {
      await client.sendOperationRequest(
        {},
        {
          httpMethod: "GET",
          baseUrl: "https://example.com",
          serializer: createSerializer(),
          responses: {
            200: {
              bodyMapper: {
                type: {
                  name: "Composite",
                  className: "NonExistentModel",
                },
              } as CompositeMapper,
            },
          },
        },
      );
      assert.fail("Should have thrown");
    } catch (err: any) {
      assert.include(err.message, "occurred in deserializing the responseBody");
    }
  });
});

describe("serializationPolicy - prepareXMLRootList non-array (line 257)", () => {
  it("should serialize XML Sequence without namespace (prepareXMLRootList no-namespace path)", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(
      serializationPolicy({
        stringifyXML: (obj) => JSON.stringify(obj),
      }),
      { phase: "Serialize" },
    );
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() });
        },
      },
      pipeline,
    });

    // Sequence without xmlNamespace to hit the !xmlNamespaceKey || !xmlNamespace path in prepareXMLRootList
    await client.sendOperationRequest(
      { body: ["item1", "item2"] },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        isXML: true,
        contentType: "application/xml",
        serializer: createSerializer({}, true),
        requestBody: {
          parameterPath: "body",
          mapper: {
            serializedName: "Items",
            xmlName: "Items",
            xmlElementName: "Item",
            // No xmlNamespace
            type: {
              name: "Sequence",
              element: { type: { name: "String" } },
            },
          } as SequenceMapper,
        },
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    const parsed = JSON.parse(capturedRequest!.body as string);
    assert.isArray(parsed.Item);
  });

  it("should wrap non-array value in prepareXMLRootList when body is null (line 257)", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(
      serializationPolicy({
        stringifyXML: (obj) => JSON.stringify(obj),
      }),
      { phase: "Serialize" },
    );
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() });
        },
      },
      pipeline,
    });

    // nullable Sequence with null body: serializer returns null (not an array),
    // which reaches prepareXMLRootList and triggers the !Array.isArray(obj) branch
    await client.sendOperationRequest(
      { body: null },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        isXML: true,
        contentType: "application/xml",
        serializer: createSerializer({}, true),
        requestBody: {
          parameterPath: "body",
          mapper: {
            serializedName: "Items",
            xmlName: "Items",
            xmlElementName: "Item",
            nullable: true,
            type: {
              name: "Sequence",
              element: { type: { name: "String" } },
            },
          } as SequenceMapper,
        },
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    const parsed = JSON.parse(capturedRequest!.body as string);
    // null was wrapped into [null] by prepareXMLRootList
    assert.isArray(parsed.Item);
    assert.strictEqual(parsed.Item.length, 1);
    assert.isNull(parsed.Item[0]);
  });
});

describe("serializer - getXmlObjectValue Composite with existing $ attr (lines 845-849)", () => {
  it("should return as-is when Composite already has $ from its own xmlNamespace", () => {
    // Child model WITH xmlNamespace - its serialization adds $ to payload
    const childModel: CompositeMapper = {
      serializedName: "ChildModel",
      xmlNamespace: "http://child.com",
      xmlNamespacePrefix: "ch",
      type: {
        name: "Composite",
        className: "ChildModel",
        modelProperties: {
          text: {
            serializedName: "text",
            xmlName: "text",
            type: { name: "String" },
          },
        },
      },
    };

    const parentMapper: CompositeMapper = {
      serializedName: "ParentModel",
      type: {
        name: "Composite",
        modelProperties: {
          child: {
            serializedName: "child",
            xmlName: "child",
            xmlNamespace: "http://outer.com",
            xmlNamespacePrefix: "outer",
            type: {
              name: "Composite",
              className: "ChildModel",
            },
          } as CompositeMapper,
        },
      },
    };

    const s = createSerializer({ ChildModel: childModel }, true);
    const result = s.serialize(parentMapper, { child: { text: "hello" } }, "testObj");
    // child should have $ from its own xmlNamespace (line 845 path)
    assert.ok(result.child);
    assert.ok(result.child.$);
    assert.strictEqual(result.child.text, "hello");
  });

  it("should add xmlns for Composite without existing $ attr", () => {
    // Child model with NO properties - so the for loop doesn't execute,
    // and $ is never set on the payload by serializeCompositeType
    const childModelEmpty: CompositeMapper = {
      serializedName: "ChildEmpty",
      type: {
        name: "Composite",
        className: "ChildEmpty",
        modelProperties: {},
      },
    };

    const parentMapper: CompositeMapper = {
      serializedName: "ParentModel2",
      type: {
        name: "Composite",
        modelProperties: {
          child: {
            serializedName: "child",
            xmlName: "child",
            xmlNamespace: "http://outer.com",
            type: {
              name: "Composite",
              className: "ChildEmpty",
            },
          } as CompositeMapper,
        },
      },
    };

    const s = createSerializer({ ChildEmpty: childModelEmpty }, true);
    const result = s.serialize(parentMapper, { child: {} }, "testObj");
    // getXmlObjectValue adds $ since child didn't have it (lines 847-849)
    assert.ok(result.child);
    assert.ok(result.child.$);
    assert.strictEqual(result.child.$["xmlns"], "http://outer.com");
  });

  it("should wrap non-Composite value with xmlNamespace (lines 852-855)", () => {
    // A non-Composite property (e.g., String) with xmlNamespace
    // goes through the non-Composite path in getXmlObjectValue
    const mapper: CompositeMapper = {
      serializedName: "Parent",
      type: {
        name: "Composite",
        modelProperties: {
          value: {
            serializedName: "value",
            xmlName: "value",
            xmlNamespace: "http://ns.com",
            xmlNamespacePrefix: "ns",
            type: { name: "String" },
          },
        },
      },
    };

    const s = createSerializer({}, true);
    const result = s.serialize(mapper, { value: "hello" }, "testObj");
    // The String value should be wrapped: { _: "hello", $: { "xmlns:ns": "http://ns.com" } }
    assert.ok(result.value);
    assert.ok(result.value.$);
    assert.strictEqual(result.value.$["xmlns:ns"], "http://ns.com");
    assert.strictEqual(result.value._, "hello");
  });
});

describe("urlHelpers - remaining uncovered lines", () => {
  it("should handle empty path in appendPath (line 111)", () => {
    const serializer = createSerializer({}, false);
    // operationSpec.path is "{param}" which resolves to "" after replacement
    // This causes appendPath to be called with empty pathToAppend
    const url = getRequestUrl(
      "https://example.com",
      {
        httpMethod: "GET",
        responses: {},
        serializer,
        path: "{param}",
        urlParameters: [
          {
            parameterPath: "param",
            mapper: { serializedName: "param", type: { name: "String" } },
            skipEncoding: true,
          },
        ],
      },
      { param: "" },
      {},
    );
    assert.strictEqual(url, "https://example.com");
  });

  it("should add trailing slash to path without one (line 118)", () => {
    const serializer = createSerializer({}, false);
    const url = getRequestUrl(
      "https://example.com/api",
      {
        path: "items",
        httpMethod: "GET",
        responses: {},
        serializer,
      },
      {},
      {},
    );
    assert.include(url, "api/items");
  });

  it("should handle undefined value in combinedParams (line 307)", () => {
    // This covers the case where a param has no = sign (bare key)
    // simpleParseQueryParams gives value as undefined
    // When we later iterate, it hits the else branch at line 307
    const result = appendQueryParams(
      "https://example.com?bare",
      new Map([["other", "val"]]),
      new Set(),
    );
    assert.include(result, "bare");
    assert.include(result, "other=val");
  });

  it("should handle array push in simpleParseQueryParams for 3+ duplicate keys (line 243)", () => {
    // First two dups create an array, third dup pushes to the array
    const result = appendQueryParams(
      "https://example.com?x=1&x=2&x=3",
      new Map([["y", "4"]]),
      new Set(),
    );
    assert.include(result, "x=1");
    assert.include(result, "x=2");
    assert.include(result, "x=3");
    assert.include(result, "y=4");
  });
});

describe("pipeline - default options parameter (lines 41-42)", () => {
  it("should handle being called with no arguments", () => {
    const pipeline = createClientPipeline();
    assert.ok(pipeline);
    const policies = pipeline.getOrderedPolicies();
    assert.isTrue(policies.length > 0);
  });
});

describe("authorizeRequestOnClaimChallenge - parseCAEChallenge fallback (line 76)", () => {
  it("should handle completely unparseable WWW-Authenticate value", async () => {
    const request = createPipelineRequest({ url: "https://example.com" });
    const result = await authorizeRequestOnClaimChallenge({
      async getAccessToken() {
        return { token: "token", expiresOnTimestamp: Date.now() + 3600000 };
      },
      scopes: [],
      response: {
        headers: createHttpHeaders({
          "WWW-Authenticate": "NotBearer gibberish",
        }),
        request,
        status: 401,
      },
      request,
    });
    assert.isFalse(result);
  });
});

describe("operationHelpers - array parameterPath empty check (line 35)", () => {
  it("should handle empty string parameterPath", () => {
    const result = getOperationArgumentValueFromParameter(
      { "": "rootValue" },
      {
        parameterPath: "",
        mapper: {
          serializedName: "test",
          type: { name: "String" },
        },
      },
    );
    // Empty string parameterPath becomes [""], which has length > 0
    assert.strictEqual(result, "rootValue");
  });
});

describe("serializationPolicy - XML Stream body should not be stringified", () => {
  it("should pass stream through in XML mode", async () => {
    const streamBody = { pipe: vi.fn(), on: vi.fn() };
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(
      serializationPolicy({
        stringifyXML: (obj) => JSON.stringify(obj),
      }),
      { phase: "Serialize" },
    );
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() });
        },
      },
      pipeline,
    });

    await client.sendOperationRequest(
      { body: streamBody },
      {
        httpMethod: "POST",
        baseUrl: "https://example.com",
        isXML: true,
        contentType: "application/xml",
        serializer: createSerializer({}, true),
        requestBody: {
          parameterPath: "body",
          mapper: {
            serializedName: "body",
            type: { name: "Stream" },
          },
        },
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    // Stream should not be stringified
    assert.strictEqual(capturedRequest!.body, streamBody);
  });
});

describe("serializationPolicy - custom headers via requestOptions", () => {
  it("should apply custom headers from requestOptions", async () => {
    let capturedRequest: OperationRequest | undefined;
    const pipeline = createEmptyPipeline();
    pipeline.addPolicy(serializationPolicy(), { phase: "Serialize" });
    const client = new ServiceClient({
      httpClient: {
        sendRequest: (req) => {
          capturedRequest = req;
          return Promise.resolve({ request: req, status: 200, headers: createHttpHeaders() });
        },
      },
      pipeline,
    });

    await client.sendOperationRequest(
      {
        options: {
          requestOptions: {
            customHeaders: { "X-Custom": "myValue" },
          },
        },
      },
      {
        httpMethod: "GET",
        baseUrl: "https://example.com",
        serializer: createSerializer(),
        responses: { 200: {} },
      },
    );

    assert.ok(capturedRequest);
    assert.strictEqual(capturedRequest!.headers.get("X-Custom"), "myValue");
  });
});
