// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace OfficeCli.Core;

public enum OutputFormat
{
    Text,
    Json
}

public class ViewResult
{
    [JsonPropertyName("view")]
    public string View { get; set; } = "";
    [JsonPropertyName("content")]
    public string Content { get; set; } = "";
}

public class NodesResult
{
    [JsonPropertyName("matches")]
    public int Matches { get; set; }
    [JsonPropertyName("results")]
    public List<DocumentNode> Results { get; set; } = new();
}

public class IssuesResult
{
    [JsonPropertyName("count")]
    public int Count { get; set; }
    [JsonPropertyName("issues")]
    public List<DocumentIssue> Issues { get; set; } = new();
}

public class ErrorResult
{
    [JsonPropertyName("error")]
    public string Error { get; set; } = "";
    [JsonPropertyName("code")]
    public string? Code { get; set; }
    [JsonPropertyName("suggestion")]
    public string? Suggestion { get; set; }
    [JsonPropertyName("help")]
    public string? Help { get; set; }
    [JsonPropertyName("validValues")]
    public string[]? ValidValues { get; set; }
}

public class CliWarning
{
    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
    [JsonPropertyName("code")]
    public string? Code { get; set; }
    [JsonPropertyName("suggestion")]
    public string? Suggestion { get; set; }
}

/// <summary>
/// Thread-static context for capturing warnings during command execution in JSON mode.
/// </summary>
public static class WarningContext
{
    [ThreadStatic]
    private static List<CliWarning>? _warnings;

    public static void Begin() => _warnings = new List<CliWarning>();

    public static void Add(string message, string? code = null, string? suggestion = null)
    {
        _warnings?.Add(new CliWarning { Message = message, Code = code, Suggestion = suggestion });
    }

    public static List<CliWarning>? End()
    {
        var result = _warnings;
        _warnings = null;
        return result?.Count > 0 ? result : null;
    }

    public static bool IsActive => _warnings != null;
}

[JsonSourceGenerationOptions(
    WriteIndented = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(ViewResult))]
[JsonSerializable(typeof(NodesResult))]
[JsonSerializable(typeof(IssuesResult))]
[JsonSerializable(typeof(ErrorResult))]
[JsonSerializable(typeof(CliWarning))]
[JsonSerializable(typeof(List<CliWarning>))]
[JsonSerializable(typeof(string[]))]
[JsonSerializable(typeof(DocumentNode))]
[JsonSerializable(typeof(List<DocumentNode>))]
[JsonSerializable(typeof(List<DocumentIssue>))]
[JsonSerializable(typeof(Dictionary<string, object?>))]
[JsonSerializable(typeof(bool))]
[JsonSerializable(typeof(int))]
[JsonSerializable(typeof(long))]
[JsonSerializable(typeof(short))]
[JsonSerializable(typeof(uint))]
[JsonSerializable(typeof(double))]
[JsonSerializable(typeof(string))]
internal partial class AppJsonContext : JsonSerializerContext;

public static class OutputFormatter
{
    public static readonly JsonSerializerOptions PublicJsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        TypeInfoResolver = AppJsonContext.Default
    };

    /// <summary>
    /// Wraps pre-serialized data JSON into a unified envelope with optional warnings.
    /// Output: { "success": true, "data": ..., "warnings": [...] }
    /// </summary>
    public static string WrapEnvelope(string dataJson, List<CliWarning>? warnings = null)
    {
        var envelope = new JsonObject { ["success"] = true };

        // Parse and embed data as-is (preserves original structure)
        try { envelope["data"] = JsonNode.Parse(dataJson); }
        catch { envelope["data"] = dataJson; } // fallback: plain string

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    /// <summary>
    /// Wraps a plain text result (like "Updated ..." or "Added ...") into an envelope.
    /// </summary>
    public static string WrapEnvelopeText(string message, List<CliWarning>? warnings = null)
    {
        var envelope = new JsonObject
        {
            ["success"] = true,
            ["message"] = message
        };

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    /// <summary>
    /// Wraps a plain text result together with the updated node state into an envelope.
    /// The node JSON is embedded as raw JSON (not double-escaped).
    /// Output: { "success": true, "message": "...", "node": {...}, "warnings": [...] }
    /// </summary>
    public static string WrapEnvelopeWithNode(string message, DocumentNode? node, List<CliWarning>? warnings = null)
    {
        var envelope = new JsonObject
        {
            ["success"] = true,
            ["message"] = message
        };

        if (node != null)
        {
            var nodeJson = FormatNode(node, OutputFormat.Json);
            try { envelope["node"] = JsonNode.Parse(nodeJson); }
            catch { envelope["node"] = nodeJson; }
        }

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    /// <summary>
    /// Wraps a failed text result (e.g. all properties unsupported) into an envelope.
    /// Output: { "success": false, "message": "...", "warnings": [...] }
    /// </summary>
    public static string WrapEnvelopeError(string message, List<CliWarning>? warnings = null)
    {
        var envelope = new JsonObject
        {
            ["success"] = false,
            ["message"] = message
        };

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    /// <summary>
    /// Wraps an error into an envelope.
    /// Output: { "success": false, "error": { ... } }
    /// </summary>
    public static string WrapErrorEnvelope(Exception ex)
    {
        var errorResult = BuildErrorResult(ex);
        var errorNode = JsonSerializer.SerializeToNode(errorResult, AppJsonContext.Default.ErrorResult)!.AsObject();

        // Embed correctedArgs if present on CliException
        if (ex is CliException { CorrectedArgs: { } correctedArgs })
        {
            var argsNode = new JsonObject();
            foreach (var (key, value) in correctedArgs)
            {
                argsNode[key] = value switch
                {
                    string s => JsonValue.Create(s),
                    string[] arr => new JsonArray(arr.Select(a => (JsonNode)JsonValue.Create(a)!).ToArray()),
                    int i => JsonValue.Create(i),
                    bool b => JsonValue.Create(b),
                    _ => JsonValue.Create(value?.ToString())
                };
            }
            errorNode["correctedArgs"] = argsNode;
        }

        var envelope = new JsonObject
        {
            ["success"] = false,
            ["error"] = errorNode
        };
        return envelope.ToJsonString(JsonOptions);
    }

    public static string FormatError(Exception ex)
    {
        return JsonSerializer.Serialize(BuildErrorResult(ex), AppJsonContext.Default.ErrorResult);
    }

    private static ErrorResult BuildErrorResult(Exception ex)
    {
        var result = new ErrorResult { Error = ex.Message };

        if (ex is CliException cli)
        {
            result.Code = cli.Code;
            result.Suggestion = cli.Suggestion;
            result.Help = cli.Help;
            result.ValidValues = cli.ValidValues;
        }
        else
        {
            EnrichFromMessage(result, ex);
        }

        return result;
    }

    private static void EnrichFromMessage(ErrorResult result, Exception ex)
    {
        var msg = ex.Message;

        // Pattern: "Slide 50 not found (total: 8)" → code=not_found, suggestion about valid range
        var notFoundMatch = System.Text.RegularExpressions.Regex.Match(msg, @"^(\w+)\s+(\d+)\s+not found \(total:\s*(\d+)\)");
        if (notFoundMatch.Success)
        {
            var elementType = notFoundMatch.Groups[1].Value;
            var total = int.Parse(notFoundMatch.Groups[3].Value);
            result.Code = "not_found";
            result.Suggestion = total == 0
                ? $"No {elementType} elements exist. Add one first."
                : $"Valid {elementType} index range: 1-{total}";
            return;
        }

        // Pattern: "Unknown part: X. Available: ..."
        var unknownPartMatch = System.Text.RegularExpressions.Regex.Match(msg, @"Unknown part: (.+?)\. Available: (.+)");
        if (unknownPartMatch.Success)
        {
            result.Code = "invalid_path";
            result.ValidValues = unknownPartMatch.Groups[2].Value.Split(", ");
            return;
        }

        // Pattern: "Unsupported file type: .xyz. Supported: ..."
        if (msg.Contains("Unsupported file type"))
        {
            result.Code = "unsupported_type";
            return;
        }

        // Pattern: "Invalid font size: ..." / "Invalid color value: ..." / "Invalid ... value"
        if (msg.StartsWith("Invalid "))
        {
            result.Code = "invalid_value";
            // Extract "Valid values: ..." if present
            var validMatch = System.Text.RegularExpressions.Regex.Match(msg, @"Valid values?:\s*(.+?)\.?$");
            if (validMatch.Success)
                result.ValidValues = validMatch.Groups[1].Value.Split(", ");
            return;
        }

        // Pattern: "UNSUPPORTED props: ..."
        if (msg.StartsWith("UNSUPPORTED props:"))
        {
            result.Code = "unsupported_property";
            result.Help = "officecli help <format>-set";
            return;
        }

        // Pattern: "'X' property is required for Y type"
        if (msg.Contains("property is required"))
        {
            result.Code = "missing_property";
            return;
        }

        // Pattern: "File not found: ..."
        if (ex is FileNotFoundException)
        {
            result.Code = "file_not_found";
            return;
        }
    }

    public static string FormatView(string view, string content, OutputFormat format)
    {
        return format switch
        {
            OutputFormat.Json => JsonSerializer.Serialize(new ViewResult { View = view, Content = content }, AppJsonContext.Default.ViewResult),
            _ => content
        };
    }

    public static string FormatNode(DocumentNode node, OutputFormat format)
    {
        if (format == OutputFormat.Json)
            return JsonSerializer.Serialize(node, AppJsonContext.Default.DocumentNode);

        return FormatNodeAsText(node, 0);
    }

    public static string FormatNodes(List<DocumentNode> nodes, OutputFormat format)
    {
        if (format == OutputFormat.Json)
            return JsonSerializer.Serialize(new NodesResult { Matches = nodes.Count, Results = nodes }, AppJsonContext.Default.NodesResult);

        var sb = new StringBuilder();
        sb.AppendLine($"Matches: {nodes.Count}");
        foreach (var node in nodes)
        {
            sb.AppendLine($"  {node.Path}: {node.Text ?? node.Preview ?? node.Type}");
            foreach (var (key, val) in node.Format)
                sb.AppendLine($"    {key}: {val}");
        }
        return sb.ToString().TrimEnd();
    }

    public static string FormatIssues(List<DocumentIssue> issues, OutputFormat format)
    {
        if (format == OutputFormat.Json)
            return JsonSerializer.Serialize(new IssuesResult { Count = issues.Count, Issues = issues }, AppJsonContext.Default.IssuesResult);

        var sb = new StringBuilder();
        sb.AppendLine($"Found {issues.Count} issue(s):");
        sb.AppendLine();

        var grouped = issues.GroupBy(i => i.Type);
        foreach (var group in grouped)
        {
            var typeName = group.Key switch
            {
                IssueType.Format => "Format Issues",
                IssueType.Content => "Content Issues",
                IssueType.Structure => "Structure Issues",
                _ => "Other"
            };
            sb.AppendLine($"{typeName} ({group.Count()}):");

            foreach (var issue in group)
            {
                var severity = issue.Severity switch
                {
                    IssueSeverity.Error => "ERROR",
                    IssueSeverity.Warning => "WARN",
                    _ => "INFO"
                };
                sb.AppendLine($"  [{issue.Id}] {issue.Path}: {issue.Message}");
                if (issue.Context != null)
                    sb.AppendLine($"       Context: \"{issue.Context}\"");
                if (issue.Suggestion != null)
                    sb.AppendLine($"       Suggestion: {issue.Suggestion}");
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    private static string FormatNodeAsText(DocumentNode node, int indent)
    {
        var sb = new StringBuilder();
        var prefix = new string(' ', indent * 2);

        sb.Append($"{prefix}{node.Path} ({node.Type})");
        if (node.Text != null) sb.Append($" \"{Truncate(node.Text, 60)}\"");
        if (node.Style != null) sb.Append($" [{node.Style}]");
        if (node.ChildCount > 0 && node.Children.Count == 0) sb.Append($" ({node.ChildCount} children)");
        sb.AppendLine();

        foreach (var (key, val) in node.Format)
            sb.AppendLine($"{prefix}  {key}: {val}");

        foreach (var child in node.Children)
            sb.Append(FormatNodeAsText(child, indent + 1));

        return sb.ToString();
    }

    private static string Truncate(string s, int maxLen)
    {
        return s.Length > maxLen ? s[..maxLen] + "..." : s;
    }
}
