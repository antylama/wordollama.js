using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using WordOllama.Contracts;
using WordOllama.Core;

namespace WordOllama.DesktopBridge;

public sealed class WordMcpOptions
{
    public const string TokenSecretName = "WORDOLLAMA_WORD_MCP_TOKEN";
    public const string EnabledSecretName = "WORDOLLAMA_WORD_MCP_ENABLED";
    private readonly object _gate = new();
    private readonly ConcurrentDictionary<string, HostMcpToken> _hostTokens = new(StringComparer.Ordinal);
    private bool _enabled;
    private string _accessToken;

    private sealed class HostMcpToken
    {
        public HostMcpToken(string token, bool enabled)
        {
            Token = token;
            Enabled = enabled;
        }

        public string Token { get; set; }
        public bool Enabled { get; set; }
    }

    public WordMcpOptions(bool enabled, string accessToken)
    {
        _enabled = enabled;
        _accessToken = accessToken;
    }

    public bool Enabled { get { lock (_gate) return _enabled; } }
    public string AccessToken { get { lock (_gate) return _accessToken; } }
    public bool HasAnyEnabled => Enabled || _hostTokens.Values.Any(entry => entry.Enabled);

    private static string CreateToken() =>
        Convert.ToBase64String(RandomNumberGenerator.GetBytes(32))
            .TrimEnd('=')
            .Replace('+', '-')
            .Replace('/', '_');

    public string GenerateAndEnable(IMutableSecretStore secrets)
    {
        var token = CreateToken();
        secrets.Set(TokenSecretName, token);
        secrets.Set(EnabledSecretName, bool.TrueString);
        lock (_gate)
        {
            _accessToken = token;
            _enabled = true;
        }
        return token;
    }

    public (bool Enabled, bool Configured, string AccessToken) GetOrCreateHostState(string hostId, bool createIfMissing = false)
    {
        lock (_gate)
        {
            if (!_hostTokens.TryGetValue(hostId, out var entry))
            {
                if (!createIfMissing) return (false, false, string.Empty);
                entry = new HostMcpToken(CreateToken(), true);
                _hostTokens[hostId] = entry;
            }
            return (entry.Enabled, true, entry.Token);
        }
    }

    public string GenerateHostToken(string hostId)
    {
        var token = CreateToken();
        lock (_gate)
        {
            _hostTokens[hostId] = new HostMcpToken(token, true);
        }
        return token;
    }

    public void EnableHost(string hostId)
    {
        lock (_gate)
        {
            if (!_hostTokens.TryGetValue(hostId, out var entry) || entry.Token.Length < 32)
                throw new InvalidOperationException("Generate a Word MCP access token before enabling the endpoint.");
            entry.Enabled = true;
        }
    }

    public void DisableHost(string hostId)
    {
        lock (_gate)
        {
            if (_hostTokens.TryGetValue(hostId, out var entry))
                entry.Enabled = false;
        }
    }

    public bool TryMatchHostToken(string suppliedToken, out string? hostId)
    {
        var suppliedBytes = Encoding.UTF8.GetBytes(suppliedToken);
        lock (_gate)
        {
            foreach (var pair in _hostTokens)
            {
                if (!pair.Value.Enabled || string.IsNullOrEmpty(pair.Value.Token)) continue;
                var expectedBytes = Encoding.UTF8.GetBytes(pair.Value.Token);
                if (expectedBytes.Length == suppliedBytes.Length &&
                    CryptographicOperations.FixedTimeEquals(expectedBytes, suppliedBytes))
                {
                    hostId = pair.Key;
                    return true;
                }
            }
        }
        hostId = null;
        return false;
    }

    public void Enable(IMutableSecretStore secrets)
    {
        lock (_gate)
        {
            if (_accessToken.Length < 32)
                throw new InvalidOperationException("Generate a Word MCP access token before enabling the endpoint.");
            secrets.Set(EnabledSecretName, bool.TrueString);
            _enabled = true;
        }
    }

    public void Disable(IMutableSecretStore secrets)
    {
        secrets.Set(EnabledSecretName, bool.FalseString);
        lock (_gate) _enabled = false;
    }
}

public static class WordMcpEndpoint
{
    public const string Path = "/mcp/word";
    private const string DefaultProtocolVersion = "2025-03-26";
    private const string StatusToolName = "wordollama_status";
    private static readonly JsonSerializerOptions WebJson = new(JsonSerializerDefaults.Web);

    public static bool TryAuthorize(
        HttpRequest request,
        WordMcpOptions options,
        string? routeToken,
        out string? scopedHostId)
    {
        scopedHostId = null;
        string? supplied = null;
        var authorization = request.Headers.Authorization.FirstOrDefault();
        if (authorization is not null && authorization.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            supplied = authorization["Bearer ".Length..].Trim();
        }
        else if (!string.IsNullOrWhiteSpace(routeToken))
        {
            supplied = routeToken.Trim();
        }
        if (string.IsNullOrEmpty(supplied)) return false;

        if (options.TryMatchHostToken(supplied, out scopedHostId))
        {
            return true;
        }

        if (!options.Enabled || string.IsNullOrWhiteSpace(options.AccessToken)) return false;
        var expectedBytes = Encoding.UTF8.GetBytes(options.AccessToken);
        var suppliedBytes = Encoding.UTF8.GetBytes(supplied);
        return expectedBytes.Length == suppliedBytes.Length &&
            CryptographicOperations.FixedTimeEquals(expectedBytes, suppliedBytes);
    }

    public static bool IsAuthorized(HttpRequest request, WordMcpOptions options, string? routeToken = null) =>
        TryAuthorize(request, options, routeToken, out _);

    private sealed class FailedAuthTracker
    {
        private const int MaxFailuresPerWindow = 15;
        private static readonly TimeSpan Window = TimeSpan.FromSeconds(60);
        private readonly object _gate = new();
        private readonly Dictionary<string, Queue<DateTimeOffset>> _failures = new(StringComparer.Ordinal);

        public bool IsRateLimited(string key, DateTimeOffset now, out int retryAfterSeconds)
        {
            lock (_gate)
            {
                if (_failures.TryGetValue(key, out var queue))
                {
                    Prune(queue, now);
                    if (queue.Count >= MaxFailuresPerWindow)
                    {
                        var oldest = queue.Peek();
                        retryAfterSeconds = Math.Max(1, (int)Math.Ceiling((oldest + Window - now).TotalSeconds));
                        return true;
                    }
                    if (queue.Count == 0) _failures.Remove(key);
                }
                retryAfterSeconds = 0;
                return false;
            }
        }

        public void RecordFailure(string key, DateTimeOffset now)
        {
            lock (_gate)
            {
                if (_failures.Count > 4096) _failures.Clear();
                if (!_failures.TryGetValue(key, out var queue))
                {
                    queue = new Queue<DateTimeOffset>();
                    _failures[key] = queue;
                }
                Prune(queue, now);
                queue.Enqueue(now);
            }
        }

        public void Clear(string key)
        {
            lock (_gate) _failures.Remove(key);
        }

        private static void Prune(Queue<DateTimeOffset> queue, DateTimeOffset now)
        {
            while (queue.Count > 0 && now - queue.Peek() >= Window)
            {
                queue.Dequeue();
            }
        }
    }

    private static readonly FailedAuthTracker AuthTracker = new();

    public static async Task<IResult> HandleAsync(
        HttpRequest request,
        OfficeToolBroker broker,
        WordMcpOptions options,
        CancellationToken cancellationToken,
        string? token = null)
    {
        if (!options.HasAnyEnabled) return Results.NotFound();

        var clientKey = request.HttpContext.Connection.RemoteIpAddress?.ToString() ?? "unknown";
        var now = DateTimeOffset.UtcNow;
        if (AuthTracker.IsRateLimited(clientKey, now, out var retryAfterSeconds))
        {
            request.HttpContext.Response.Headers.RetryAfter = retryAfterSeconds.ToString();
            return Results.StatusCode(StatusCodes.Status429TooManyRequests);
        }

        if (!TryAuthorize(request, options, token, out var scopedHostId))
        {
            AuthTracker.RecordFailure(clientKey, now);
            return Results.Unauthorized();
        }
        AuthTracker.Clear(clientKey);

        JsonDocument document;
        try
        {
            document = await JsonDocument.ParseAsync(request.Body, cancellationToken: cancellationToken);
        }
        catch (JsonException)
        {
            return JsonRpcError(null, -32700, "Parse error");
        }

        using (document)
        {
            var root = document.RootElement;
            var hasId = false;
            JsonElement? id = null;
            if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("id", out var idValue))
            {
                hasId = true;
                id = idValue.Clone();
            }
            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("method", out var methodValue) ||
                methodValue.ValueKind != JsonValueKind.String)
            {
                return JsonRpcError(id, -32600, "Invalid Request");
            }

            var method = methodValue.GetString()!;
            root.TryGetProperty("params", out var parameters);
            if (!hasId)
            {
                // The server is stateless. Initialization and cancellation
                // notifications require no response body.
                return Results.StatusCode(StatusCodes.Status202Accepted);
            }

            return method switch
            {
                "initialize" => Initialize(id, parameters),
                "ping" => JsonRpcResult(id, new { }),
                "tools/list" => ListTools(id, broker, scopedHostId),
                "tools/call" => await CallToolAsync(id, parameters, broker, scopedHostId, cancellationToken),
                _ => JsonRpcError(id, -32601, $"Method not found: {method}"),
            };
        }
    }

    private static IResult Initialize(JsonElement? id, JsonElement parameters)
    {
        var requestedVersion = parameters.ValueKind == JsonValueKind.Object &&
            parameters.TryGetProperty("protocolVersion", out var version) &&
            version.ValueKind == JsonValueKind.String
                ? version.GetString()
                : null;
        return JsonRpcResult(id, new
        {
            protocolVersion = string.IsNullOrWhiteSpace(requestedVersion)
                ? DefaultProtocolVersion
                : requestedVersion,
            capabilities = new { tools = new { listChanged = false } },
            serverInfo = new { name = "WordOllama Word Host", version = "1.0" },
            instructions = "Operate through a connected WordOllama task pane. Call wordollama_status first. If more than one host is connected, pass host_id on every Word tool call. Read before writing. Write tools modify the live document and should require user approval.",
        });
    }

    private static IResult ListTools(JsonElement? id, OfficeToolBroker broker, string? scopedHostId)
    {
        var tools = new List<object>
        {
            new
            {
                name = StatusToolName,
                title = "Word host status",
                description = "List connected WordOllama task panes and their opaque host_id values. No document paths or contents are disclosed.",
                inputSchema = new { type = "object", properties = new { } },
                annotations = new
                {
                    readOnlyHint = true,
                    destructiveHint = false,
                    idempotentHint = true,
                    openWorldHint = false,
                },
            },
        };
        tools.AddRange(broker.GetTools(scopedHostId)
            .Where(tool => !IsInternalPanelOnlyTool(tool.Name))
            .Select(ToMcpTool));
        return JsonRpcResult(id, new { tools });
    }

    private static bool IsInternalPanelOnlyTool(string name) =>
        string.Equals(name, "ask_human", StringComparison.OrdinalIgnoreCase);

    private static async Task<IResult> CallToolAsync(
        JsonElement? id,
        JsonElement parameters,
        OfficeToolBroker broker,
        string? scopedHostId,
        CancellationToken cancellationToken)
    {
        if (parameters.ValueKind != JsonValueKind.Object ||
            !parameters.TryGetProperty("name", out var nameValue) ||
            nameValue.ValueKind != JsonValueKind.String)
        {
            return JsonRpcError(id, -32602, "tools/call requires a tool name");
        }

        var name = nameValue.GetString()!;
        if (IsInternalPanelOnlyTool(name))
        {
            return JsonRpcError(id, -32601, $"Tool not available over MCP: {name}");
        }
        if (string.Equals(name, StatusToolName, StringComparison.OrdinalIgnoreCase))
        {
            return ToolResult(id, JsonSerializer.Serialize(broker.GetStatus(scopedHostId), WebJson), false);
        }

        var arguments = parameters.TryGetProperty("arguments", out var suppliedArguments) &&
            suppliedArguments.ValueKind == JsonValueKind.Object
                ? suppliedArguments.Clone()
                : JsonSerializer.SerializeToElement(new { });
        var requestedHostId = arguments.TryGetProperty("host_id", out var hostValue) &&
            hostValue.ValueKind == JsonValueKind.String
                ? hostValue.GetString()
                : null;
        var effectiveHostId = scopedHostId ?? requestedHostId;
        try
        {
            var result = await broker.CallAsync(name, arguments, effectiveHostId, cancellationToken);
            return ToolResult(id, result.Result, result.IsError);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            return ToolResult(id, exception.Message, true);
        }
    }

    private static object ToMcpTool(OfficeToolDescriptor tool) => new
    {
        name = tool.Name,
        title = tool.Name.Replace('_', ' '),
        description = tool.Description,
        inputSchema = AddHostId(tool.ParameterSchema),
        annotations = new
        {
            readOnlyHint = !tool.IsWriteOperation,
            destructiveHint = tool.IsWriteOperation,
            idempotentHint = false,
            openWorldHint = false,
        },
    };

    private static JsonElement AddHostId(JsonElement inputSchema)
    {
        var schema = JsonNode.Parse(inputSchema.GetRawText()) as JsonObject ?? new JsonObject
        {
            ["type"] = "object",
        };
        var properties = schema["properties"] as JsonObject ?? new JsonObject();
        properties["host_id"] = new JsonObject
        {
            ["type"] = "string",
            ["description"] = "Opaque host identifier from wordollama_status; required when multiple Word task panes are connected.",
        };
        schema["properties"] = properties;
        return JsonSerializer.SerializeToElement(schema);
    }

    private static IResult ToolResult(JsonElement? id, string text, bool isError) =>
        JsonRpcResult(id, new
        {
            content = new[] { new { type = "text", text } },
            isError,
        });

    private static IResult JsonRpcResult(JsonElement? id, object result) =>
        Results.Json(new { jsonrpc = "2.0", id, result });

    private static IResult JsonRpcError(JsonElement? id, int code, string message) =>
        Results.Json(new { jsonrpc = "2.0", id, error = new { code, message } });
}
