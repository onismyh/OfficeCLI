// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Minimal MCP (Model Context Protocol) server over stdio.
/// Implements JSON-RPC 2.0 with initialize, tools/list, and tools/call.
/// All JSON is hand-written via Utf8JsonWriter to avoid reflection (PublishTrimmed).
/// </summary>
public static class McpServer
{
    public static async Task RunAsync()
    {
        using var reader = new StreamReader(Console.OpenStandardInput());
        using var writer = new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true };

        while (true)
        {
            var line = await reader.ReadLineAsync();
            if (line == null) break;
            if (string.IsNullOrWhiteSpace(line)) continue;

            try
            {
                using var doc = JsonDocument.Parse(line);
                var root = doc.RootElement;
                var method = root.TryGetProperty("method", out var m) ? m.GetString() : null;
                var id = root.TryGetProperty("id", out var idEl) ? idEl.Clone() : (JsonElement?)null;

                var response = method switch
                {
                    "initialize" => HandleInitialize(id),
                    "notifications/initialized" => null,
                    "tools/list" => HandleToolsList(id),
                    "tools/call" => HandleToolsCall(id, root),
                    "ping" => WriteJson(w => { w.WriteStartObject(); Rpc(w, id); w.WriteStartObject("result"); w.WriteEndObject(); w.WriteEndObject(); }),
                    _ => id.HasValue ? ErrorJson(id, -32601, $"Method not found: {method}") : null,
                };

                if (response != null)
                    await writer.WriteLineAsync(response);
            }
            catch (JsonException)
            {
                await writer.WriteLineAsync(ErrorJson(null, -32700, "Parse error"));
            }
            catch (Exception ex)
            {
                await writer.WriteLineAsync(ErrorJson(null, -32603, $"Internal error: {ex.Message}"));
            }
        }
    }

    // ==================== Handlers ====================

    private static string HandleInitialize(JsonElement? id) => WriteJson(w =>
    {
        w.WriteStartObject();
        Rpc(w, id);
        w.WriteStartObject("result");
        w.WriteString("protocolVersion", "2024-11-05");
        w.WriteStartObject("capabilities");
        w.WriteStartObject("tools"); w.WriteBoolean("listChanged", false); w.WriteEndObject();
        w.WriteEndObject();
        w.WriteStartObject("serverInfo"); w.WriteString("name", "officecli"); w.WriteString("version", "1.0.28"); w.WriteEndObject();
        w.WriteEndObject();
        w.WriteEndObject();
    });

    private static string HandleToolsList(JsonElement? id) => WriteJson(w =>
    {
        w.WriteStartObject();
        Rpc(w, id);
        w.WriteStartObject("result");
        w.WriteStartArray("tools");
        WriteToolDefinitions(w);
        w.WriteEndArray();
        w.WriteEndObject();
        w.WriteEndObject();
    });

    private static string HandleToolsCall(JsonElement? id, JsonElement root)
    {
        if (!root.TryGetProperty("params", out var p))
            return ErrorJson(id, -32602, "Missing params");
        var name = p.TryGetProperty("name", out var n) ? n.GetString() : null;
        var args = p.TryGetProperty("arguments", out var a) ? a : default;
        if (string.IsNullOrEmpty(name))
            return ErrorJson(id, -32602, "Missing tool name");

        try
        {
            // Route by tool name:
            //   New multi-tool: officecli_view -> "view", officecli_raw_set -> "raw-set"
            //   Legacy single tool: officecli + command arg -> route by command value
            string toolName;
            if (name.StartsWith("officecli_"))
            {
                toolName = name["officecli_".Length..].Replace('_', '-');
            }
            else if (name == "officecli" && args.ValueKind == JsonValueKind.Object && args.TryGetProperty("command", out var cmd))
            {
                toolName = cmd.GetString() ?? name;
            }
            else
            {
                toolName = name;
            }
            var result = ExecuteTool(toolName, args);
            return WriteJson(w =>
            {
                w.WriteStartObject();
                Rpc(w, id);
                w.WriteStartObject("result");
                w.WriteStartArray("content");
                w.WriteStartObject(); w.WriteString("type", "text"); w.WriteString("text", result); w.WriteEndObject();
                w.WriteEndArray();
                w.WriteBoolean("isError", false);
                w.WriteEndObject();
                w.WriteEndObject();
            });
        }
        catch (Exception ex)
        {
            return WriteJson(w =>
            {
                w.WriteStartObject();
                Rpc(w, id);
                w.WriteStartObject("result");
                w.WriteStartArray("content");
                w.WriteStartObject(); w.WriteString("type", "text"); w.WriteString("text", $"Error: {ex.Message}"); w.WriteEndObject();
                w.WriteEndArray();
                w.WriteBoolean("isError", true);
                w.WriteEndObject();
                w.WriteEndObject();
            });
        }
    }

    // ==================== Tool Execution ====================

    private static string ExecuteTool(string name, JsonElement args)
    {
        string Arg(string key) => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(key, out var v) ? v.GetString() ?? "" : "";
        int ArgInt(string key, int def) => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(key, out var v) && v.TryGetInt32(out var i) ? i : def;
        int? ArgIntOpt(string key) => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(key, out var v) && v.TryGetInt32(out var i) ? i : null;
        string[] ArgStringArray(string key)
        {
            if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(key, out var v) || v.ValueKind != JsonValueKind.Array) return [];
            return v.EnumerateArray().Select(e => e.GetString() ?? "").ToArray();
        }

        switch (name)
        {
            case "create":
            {
                var file = Arg("file");
                BlankDocCreator.Create(file);
                return $"Created {file}";
            }
            case "view":
            {
                var file = Arg("file");
                var mode = Arg("mode");
                var start = ArgIntOpt("start");
                var end = ArgIntOpt("end");
                var maxLines = ArgIntOpt("max_lines");
                var rawColsStr = Arg("cols");
                var cols = !string.IsNullOrEmpty(rawColsStr)
                    ? new HashSet<string>(rawColsStr.Split(',').Select(c => c.Trim().ToUpperInvariant()))
                    : null;
                var issueType = Arg("type"); if (string.IsNullOrEmpty(issueType)) issueType = null;
                using var handler = DocumentHandlerFactory.Open(file);
                if (mode is "html" or "h")
                {
                    var html = handler.ViewAsHtml(start, end);
                    if (html != null) return html;
                }
                if (mode is "svg" or "g" && handler is Handlers.PowerPointHandler pptSvg)
                    return pptSvg.ViewAsSvg(start ?? 1);
                return mode.ToLowerInvariant() switch
                {
                    "text" or "t" => handler.ViewAsText(start, end, maxLines, cols),
                    "annotated" or "a" => handler.ViewAsAnnotated(start, end, maxLines, cols),
                    "outline" or "o" => handler.ViewAsOutline(),
                    "stats" or "s" => handler.ViewAsStats(),
                    "issues" or "i" => OutputFormatter.FormatIssues(handler.ViewAsIssues(issueType, null), OutputFormat.Json),
                    "forms" or "f" => handler is Handlers.WordHandler wfh
                        ? wfh.ViewAsFormsJson().ToJsonString(OutputFormatter.PublicJsonOptions)
                        : throw new ArgumentException("Forms view is only supported for .docx files."),
                    _ => throw new ArgumentException($"Unknown mode: {mode}")
                };
            }
            case "get":
            {
                var file = Arg("file");
                var path = Arg("path"); if (string.IsNullOrEmpty(path)) path = "/";
                var depth = ArgInt("depth", 1);
                using var handler = DocumentHandlerFactory.Open(file);
                var node = handler.Get(path, depth);
                return OutputFormatter.FormatNode(node, OutputFormat.Json);
            }
            case "query":
            {
                var file = Arg("file");
                var selector = Arg("selector");
                using var handler = DocumentHandlerFactory.Open(file);
                var filters = AttributeFilter.Parse(selector);
                var (results, _) = AttributeFilter.ApplyWithWarnings(handler.Query(selector), filters);
                return OutputFormatter.FormatNodes(results, OutputFormat.Json);
            }
            case "set":
            {
                var file = Arg("file");
                var path = Arg("path");
                var props = ParseProps(ArgStringArray("props"));
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var unsupported = handler.Set(path, props);
                var applied = props.Where(kv => !unsupported.Contains(kv.Key)).ToList();
                var msg = applied.Count > 0
                    ? $"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}"
                    : $"No properties applied to {path}";
                if (unsupported.Count > 0)
                    msg += $"\nUnsupported: {string.Join(", ", unsupported)}";
                // Echo updated node state for AI agents
                if (applied.Count > 0)
                {
                    var node = handler.Get(path);
                    msg += $"\n{OutputFormatter.FormatNode(node, OutputFormat.Json)}";
                }
                return msg;
            }
            case "add":
            {
                var file = Arg("file");
                var parent = Arg("parent");
                var type = Arg("type");
                var from = Arg("from"); if (string.IsNullOrEmpty(from)) from = null;
                var index = ArgIntOpt("index");
                var props = ParseProps(ArgStringArray("props"));
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                if (from != null)
                {
                    var resultPath = handler.CopyFrom(from, parent, index);
                    var msg = $"Copied to {resultPath}";
                    // Echo copied node state for AI agents
                    var copyNode = handler.Get(resultPath);
                    msg += $"\n{OutputFormatter.FormatNode(copyNode, OutputFormat.Json)}";
                    return msg;
                }
                else
                {
                    var resultPath = handler.Add(parent, type, index, props);
                    var msg = $"Added {type} at {resultPath}";
                    // Echo new node state for AI agents
                    var addNode = handler.Get(resultPath);
                    msg += $"\n{OutputFormatter.FormatNode(addNode, OutputFormat.Json)}";
                    return msg;
                }
            }
            case "remove":
            {
                var file = Arg("file");
                var path = Arg("path");
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                handler.Remove(path);
                return $"Removed {path}";
            }
            case "move":
            {
                var file = Arg("file");
                var path = Arg("path");
                var to = Arg("to"); if (string.IsNullOrEmpty(to)) to = null;
                var index = ArgIntOpt("index");
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var resultPath = handler.Move(path, to, index);
                var msg = $"Moved to {resultPath}";
                // Echo moved node state for AI agents
                var moveNode = handler.Get(resultPath);
                msg += $"\n{OutputFormatter.FormatNode(moveNode, OutputFormat.Json)}";
                return msg;
            }
            case "validate":
            {
                var file = Arg("file");
                using var handler = DocumentHandlerFactory.Open(file);
                var errors = handler.Validate();
                if (errors.Count == 0) return "Validation passed: no errors found.";
                var lines = errors.Select(e => $"[{e.ErrorType}] {e.Description}" +
                    (e.Path != null ? $" (Path: {e.Path})" : ""));
                return $"Found {errors.Count} error(s):\n{string.Join("\n", lines)}";
            }
            case "batch":
            {
                var file = Arg("file");
                var commands = Arg("commands");
                var items = JsonSerializer.Deserialize<List<BatchItem>>(commands, BatchJsonContext.Default.ListBatchItem);
                if (items == null || items.Count == 0)
                    throw new ArgumentException("No commands found in input.");
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var results = new List<BatchResult>();
                foreach (var item in items)
                {
                    try
                    {
                        var output = CommandBuilder.ExecuteBatchItem(handler, item, true);
                        results.Add(new BatchResult { Success = true, Output = output });
                    }
                    catch (Exception ex)
                    {
                        results.Add(new BatchResult { Success = false, Error = ex.Message });
                    }
                }
                return JsonSerializer.Serialize(results, BatchJsonContext.Default.ListBatchResult);
            }
            case "raw":
            {
                var file = Arg("file");
                var part = Arg("part"); if (string.IsNullOrEmpty(part)) part = "/document";
                var startRow = ArgIntOpt("start");
                var endRow = ArgIntOpt("end");
                var rawColsStr = Arg("cols");
                var rawCols = !string.IsNullOrEmpty(rawColsStr)
                    ? new HashSet<string>(rawColsStr.Split(',').Select(c => c.Trim().ToUpperInvariant()))
                    : null;
                using var handler = DocumentHandlerFactory.Open(file);
                return handler.Raw(part, startRow, endRow, rawCols);
            }
            case "raw-set":
            {
                var file = Arg("file");
                var part = Arg("part"); if (string.IsNullOrEmpty(part)) part = "/document";
                var xpath = Arg("xpath");
                var action = Arg("action");
                var xml = Arg("xml"); if (string.IsNullOrEmpty(xml)) xml = null;
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                handler.RawSet(part, xpath, action, xml);
                return $"raw-set applied: {action} at {xpath}";
            }
            case "schema":
            {
                var format = Arg("format").ToLowerInvariant();
                var element = Arg("element"); if (string.IsNullOrEmpty(element)) element = null;
                return SchemaCommand.Execute(format, element);
            }
            case "help":
            {
                var format = Arg("format").ToLowerInvariant();
                const string strategy = @"## Strategy
Use view (outline/stats/issues) to understand the document first, then get/query to inspect details, then set/add/remove to modify.
For 3+ mutations on the same file, use batch (one open/save cycle) instead of separate calls.
Get output keys can be used directly as Set input keys (round-trip safe).
Colors: FF0000, red, rgb(255,0,0), accent1. Sizes: 24pt. Positions: 2cm, 1in, 72pt, or raw EMU.

";
                var reference = format switch
                {
                    "xlsx" => @"# XLSX Reference

## Add types
sheet, row, cell, col, run (rich text in cell), shape, chart, picture, comment, namedrange, table, validation, pivottable, autofilter, pagebreak, colbreak
cf (conditional formatting): set type= to databar|colorscale|iconset|formula|topn|aboveaverage|duplicatevalues|uniquevalues|containstext|dateoccurring

## Cell properties (Set/Add)
value, formula, arrayformula, type (string|number|boolean), clear, link
bold, italic, strike, underline (true|single|double), superscript, subscript
font.color (#FF0000), font.size (14pt), font.name (Calibri), fill (#4472C4)
border.all (thin|medium|thick), border.left/right/top/bottom, border.color
alignment.horizontal (left|center|right), alignment.vertical, alignment.wrapText
numfmt (0%|#,##0.00|...), rotation (0-180), indent, shrinktofit
locked (true|false), formulahidden (true|false)

## Sheet properties (Set)
name, freeze (A2|B3|none), zoom (75-200), tabcolor (#FF0000|none)
autofilter (A1:F100|none), merge (A1:D1), protect (true|false), password
printarea ($A$1:$D$10|none), orientation (landscape|portrait), papersize (1=Letter|9=A4)
fittopage (1x2|true), header (&CPage &P), footer (&LConfidential), sort (A:asc,B:desc|none)

## Run properties (Set /Sheet/A1/run[N])
text, bold, italic, strike, underline, superscript, subscript, size, color, font

## CF properties
sqref/range, color (font), fill, bold, italic, strike, underline, border (thin|medium), numfmt
topn: rank, bottom (true), percent (true)
aboveaverage: below (true)
containstext: text
dateoccurring: period (today|yesterday|tomorrow|last7days|thisweek|lastweek|thismonth|lastmonth)",

                    "pptx" => @"# PPTX Reference

## Add types
slide, shape, textbox, picture, chart, table, row, cell, paragraph, run
group, connector, animation, video, equation, notes, zoom

## Shape properties (Set/Add)
text, bold, italic, underline, strike, superscript, subscript
color (#FF0000), font (Arial), size (24pt), align (left|center|right)
fill (#4472C4|gradient), outline (#000000), rotation (45)
x, y, width, height (in cm/in/pt/emu)
shadow, glow, reflection, softedge, effect3d
link (https://...), alt (alt text)

## Slide properties (Set)
layout, background, transition, notes",

                    "docx" => @"# DOCX Reference

## Add types
paragraph, run, table, row, cell, picture, hyperlink, section
style, chart, equation, footnote, endnote, bookmark, comment
toc, pagebreak, header, footer, watermark, sdt

## Run properties (Set/Add)
text, bold, italic, underline, strike, superscript, subscript
color (#FF0000), font (Arial), size (14pt), highlight
caps, smallcaps, vanish

## Paragraph properties (Set/Add)
alignment (left|center|right|justify)
spaceBefore (12pt), spaceAfter (6pt), lineSpacing (1.5x|18pt)
indent, hanging, firstline
pagebreakbefore (true|false)

## Section properties
pagewidth, pageheight, orientation (landscape|portrait)
margintop, marginbottom, marginleft, marginright",

                    _ => null
                };
                if (reference == null)
                    return "Supported formats: xlsx, pptx, docx. Call help with one of these.";
                return strategy + reference;
            }
            default:
                throw new ArgumentException($"Unknown tool: {name}");
        }
    }

    private static Dictionary<string, string> ParseProps(string[] propStrs)
    {
        var props = new Dictionary<string, string>();
        foreach (var p in propStrs)
        {
            var eq = p.IndexOf('=');
            if (eq > 0) props[p[..eq]] = p[(eq + 1)..];
        }
        return props;
    }

    // ==================== Tool Definitions ====================

    private static void WriteToolDefinitions(Utf8JsonWriter w)
    {
        // officecli_create
        WriteToolDef(w, "officecli_create",
            "Create a blank Office document. Supported extensions: .docx, .xlsx, .pptx.",
            props => { Prop(props, "file", "string", "Document file path to create (e.g. report.pptx)"); },
            required: ["file"]);

        // officecli_view
        WriteToolDef(w, "officecli_view",
            "View document content in various modes. Paths are 1-based. Modes: text (plain text with element paths), annotated (text with formatting details), outline (heading structure), stats (counts and styles), issues (formatting/content problems), html (rendered HTML), svg (PPT slide as SVG), forms (DOCX form fields).",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                PropEnum(props, "mode", "string", "View mode", ["text", "annotated", "outline", "stats", "issues", "html", "svg", "forms"]);
                Prop(props, "start", "number", "Start at paragraph/slide/row number");
                Prop(props, "end", "number", "End at paragraph/slide/row number");
                Prop(props, "max_lines", "number", "Limit output lines (shows total count when truncated)");
                Prop(props, "cols", "string", "Column filter for Excel, comma-separated (e.g. A,B,C)");
                Prop(props, "type", "string", "Issue filter: format, content, structure (issues mode only)");
            },
            required: ["file", "mode"]);

        // officecli_get
        WriteToolDef(w, "officecli_get",
            "Get a document node by DOM path. Returns type, properties, and children. Paths are 1-based: /slide[1]/shape[2], /body/p[3]/r[1], /Sheet1/A1, /header[1]. Use depth to control child traversal.",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                Prop(props, "path", "string", "DOM path (e.g. /slide[1]/shape[1], /Sheet1/A1, /body/p[1]). Defaults to /");
                Prop(props, "depth", "number", "Depth of child nodes to include (default 1)");
            },
            required: ["file"]);

        // officecli_query
        WriteToolDef(w, "officecli_query",
            "Query elements with CSS-like selectors. Supports: element types (paragraph, run, table, picture, shape), attribute filters [attr=value], pseudo-selectors (:contains, :empty, :no-alt), child combinator (>). Example: paragraph[style=Heading1], run[bold=true], shape[fill~=blue].",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                Prop(props, "selector", "string", "CSS-like selector (e.g. run[bold=true], paragraph:contains(\"hello\"))");
                Prop(props, "text", "string", "Additional text filter on results");
            },
            required: ["file", "selector"]);

        // officecli_set
        WriteToolDef(w, "officecli_set",
            "Set properties on a document element. Props are key=value strings. Common props: text, bold, italic, font, size, color (#FF0000), fill, width, height, x, y (with units: cm, in, pt, emu). Call officecli_help or officecli_schema for full property reference.",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                Prop(props, "path", "string", "DOM path to target element (e.g. /slide[1]/shape[2], /body/p[1]/r[1], /Sheet1/A1)");
                PropArray(props, "props", "key=value property pairs (e.g. [\"bold=true\", \"color=FF0000\", \"text=Hello\"])");
                Prop(props, "force", "boolean", "Force write even if document is protected");
            },
            required: ["file", "path", "props"]);

        // officecli_add
        WriteToolDef(w, "officecli_add",
            "Add a new element to the document. Requires 'type' for new elements, or 'from' to clone an existing element. Types vary by format: slide, shape, textbox, picture, table, row, cell, paragraph, run, chart, etc. Props are key=value strings for initial properties.",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                Prop(props, "parent", "string", "Parent DOM path (e.g. /slide[1], /body, /Sheet1)");
                Prop(props, "type", "string", "Element type to add (slide, shape, textbox, picture, paragraph, run, table, row, cell, chart, etc.). Required unless using 'from'.");
                Prop(props, "from", "string", "DOM path of element to clone (e.g. /slide[1]/shape[2], /body/p[1]). Cross-part relationships handled automatically.");
                Prop(props, "index", "number", "Insert position (0-based). Omit to append at end");
                PropArray(props, "props", "key=value property pairs for the new element (e.g. [\"text=Hello\", \"bold=true\"])");
                Prop(props, "force", "boolean", "Force write even if document is protected");
            },
            required: ["file", "parent"]);

        // officecli_remove
        WriteToolDef(w, "officecli_remove",
            "Remove an element from the document by DOM path. Paths are 1-based: /slide[1]/shape[2], /body/p[3], /Sheet1/A1.",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                Prop(props, "path", "string", "DOM path of element to remove (e.g. /slide[1]/shape[2], /body/p[3])");
            },
            required: ["file", "path"]);

        // officecli_move
        WriteToolDef(w, "officecli_move",
            "Move an element to a new position or parent. Specify 'to' for a new parent path, 'index' for position within the parent.",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                Prop(props, "path", "string", "DOM path of element to move (e.g. /slide[1]/shape[2])");
                Prop(props, "to", "string", "Target parent path (omit to reorder within current parent)");
                Prop(props, "index", "number", "Target position (0-based) within the parent");
            },
            required: ["file", "path"]);

        // officecli_batch
        WriteToolDef(w, "officecli_batch",
            "Execute multiple commands in a single open/save cycle. Use for 3+ mutations on the same file. Commands is a JSON array of objects with: command (get|set|add|remove|move|view|raw|validate), path, props (object), parent, type, index, etc.",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                Prop(props, "commands", "string", "JSON array of batch commands, e.g. [{\"command\":\"set\",\"path\":\"/Sheet1/A1\",\"props\":{\"value\":\"hello\"}}]");
            },
            required: ["file", "commands"]);

        // officecli_validate
        WriteToolDef(w, "officecli_validate",
            "Validate an Office document for structural errors and formatting issues.",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
            },
            required: ["file"]);

        // officecli_raw
        WriteToolDef(w, "officecli_raw",
            "View raw OpenXML of a document part. Parts: /document, /styles, /header[N], /footer[N], /slide[N], /SheetName, /theme, /numbering. Use for advanced inspection.",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                Prop(props, "part", "string", "Part path (e.g. /document, /styles, /slide[1]). Defaults to /document");
                Prop(props, "start", "number", "Start row number (Excel sheets only)");
                Prop(props, "end", "number", "End row number (Excel sheets only)");
                Prop(props, "cols", "string", "Column filter, comma-separated (Excel only, e.g. A,B,C)");
            },
            required: ["file"]);

        // officecli_raw_set
        WriteToolDef(w, "officecli_raw_set",
            "Modify raw OpenXML in a document part. Universal fallback for any OpenXML operation not covered by set/add. Actions: append, prepend, insertbefore, insertafter, replace, remove, setattr.",
            props =>
            {
                Prop(props, "file", "string", "Document file path (.docx, .xlsx, .pptx)");
                Prop(props, "part", "string", "Part path (e.g. /document, /styles, /slide[1])");
                Prop(props, "xpath", "string", "XPath to target element(s)");
                Prop(props, "action", "string", "Action: append, prepend, insertbefore, insertafter, replace, remove, setattr");
                Prop(props, "xml", "string", "XML fragment or attr=value for setattr");
            },
            required: ["file", "part", "xpath", "action"]);

        // officecli_help
        WriteToolDef(w, "officecli_help",
            "Get detailed property reference for a document format. Returns strategy guide, element types, and property reference. Call this before using set/add to know available properties.",
            props =>
            {
                PropEnum(props, "format", "string", "Document format", ["xlsx", "pptx", "docx"]);
            },
            required: ["format"]);

        // officecli_schema
        WriteToolDef(w, "officecli_schema",
            "Get structured property definitions for a document format. Returns JSON with element types and their settable properties, types, and examples. Use this for programmatic discovery of available properties.",
            props =>
            {
                PropEnum(props, "format", "string", "Document format", ["docx", "xlsx", "pptx"]);
                Prop(props, "element", "string", "Optional element type (e.g. shape, paragraph, cell, slide) to get detailed property definitions");
            },
            required: ["format"]);
    }

    // ---- Helpers for writing tool definitions ----

    private static void WriteToolDef(Utf8JsonWriter w, string name, string description,
        Action<Utf8JsonWriter> writeProperties, string[] required)
    {
        w.WriteStartObject();
        w.WriteString("name", name);
        w.WriteString("description", description);
        w.WriteStartObject("inputSchema");
        w.WriteString("type", "object");
        w.WriteStartObject("properties");
        writeProperties(w);
        w.WriteEndObject(); // end properties
        w.WriteStartArray("required");
        foreach (var r in required) w.WriteStringValue(r);
        w.WriteEndArray();
        w.WriteEndObject(); // end inputSchema
        w.WriteEndObject(); // end tool
    }

    private static void Prop(Utf8JsonWriter w, string name, string type, string description)
    {
        w.WriteStartObject(name);
        w.WriteString("type", type);
        w.WriteString("description", description);
        w.WriteEndObject();
    }

    private static void PropEnum(Utf8JsonWriter w, string name, string type, string description, string[] values)
    {
        w.WriteStartObject(name);
        w.WriteString("type", type);
        w.WriteStartArray("enum");
        foreach (var v in values) w.WriteStringValue(v);
        w.WriteEndArray();
        w.WriteString("description", description);
        w.WriteEndObject();
    }

    private static void PropArray(Utf8JsonWriter w, string name, string description)
    {
        w.WriteStartObject(name);
        w.WriteString("type", "array");
        w.WriteStartObject("items"); w.WriteString("type", "string"); w.WriteEndObject();
        w.WriteString("description", description);
        w.WriteEndObject();
    }

    // ==================== JSON-RPC Helpers ====================

    private static string WriteJson(Action<Utf8JsonWriter> build)
    {
        using var ms = new MemoryStream();
        using (var w = new Utf8JsonWriter(ms)) build(w);
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    private static void Rpc(Utf8JsonWriter w, JsonElement? id)
    {
        w.WriteString("jsonrpc", "2.0");
        if (id.HasValue) { w.WritePropertyName("id"); id.Value.WriteTo(w); }
        else w.WriteNull("id");
    }

    private static string ErrorJson(JsonElement? id, int code, string message) => WriteJson(w =>
    {
        w.WriteStartObject();
        Rpc(w, id);
        w.WriteStartObject("error");
        w.WriteNumber("code", code);
        w.WriteString("message", message);
        w.WriteEndObject();
        w.WriteEndObject();
    });
}
