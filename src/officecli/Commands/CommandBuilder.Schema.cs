// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command CreateSchemaCommand()
    {
        var formatArg = new Argument<string>("format") { Description = "Document format: docx, xlsx, pptx" };
        var elementArg = new Argument<string?>("element") { Description = "Element type (e.g. paragraph, shape, cell, slide)" };
        elementArg.DefaultValueFactory = _ => null;

        var schemaCommand = new Command("schema", "Show structured property definitions for a document format (JSON output for AI agents)")
        {
            formatArg,
            elementArg
        };

        schemaCommand.SetAction(result =>
        {
            var format = result.GetValue(formatArg)!.ToLowerInvariant();
            var element = result.GetValue(elementArg);
            var output = SchemaCommand.Execute(format, element);
            Console.WriteLine(output);
            // Return non-zero exit code if the output is an error response
            return SchemaCommand.IsError(output) ? 1 : 0;
        });

        return schemaCommand;
    }
}

/// <summary>
/// Generates structured JSON schema definitions describing settable elements and properties
/// for each Office document format. Designed for AI agent self-discovery.
/// </summary>
internal static class SchemaCommand
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static string Execute(string format, string? element)
    {
        if (format is not ("docx" or "xlsx" or "pptx"))
            return MakeError("Unsupported format. Use: docx, xlsx, pptx");

        if (string.IsNullOrEmpty(element))
            return FormatOverview(format);

        return ElementDetail(format, element!.ToLowerInvariant());
    }

    /// <summary>Check if the output is an error response.</summary>
    public static bool IsError(string output) => output.TrimStart().StartsWith("{") &&
        output.Contains("\"error\":");

    private static string MakeError(string message) => new JsonObject
    {
        ["error"] = message
    }.ToJsonString(JsonOptions);

    // ==================== Format overview ====================

    private static string FormatOverview(string format)
    {
        var obj = new JsonObject { ["format"] = format };

        var (elements, commonProps) = format switch
        {
            "pptx" => (
                new JsonArray("slide", "shape", "textbox", "picture", "table", "row", "cell",
                    "paragraph", "run", "chart", "group", "connector", "animation", "video",
                    "equation", "notes", "zoom", "placeholder"),
                new JsonArray("text", "bold", "italic", "underline", "strike", "font", "size",
                    "color", "fill", "x", "y", "width", "height", "rotation", "opacity",
                    "align", "link", "alt", "shadow", "glow", "line", "gradient")
            ),
            "docx" => (
                new JsonArray("paragraph", "run", "table", "row", "cell", "picture",
                    "hyperlink", "section", "style", "chart", "equation", "footnote",
                    "endnote", "bookmark", "comment", "toc", "pagebreak", "header",
                    "footer", "watermark", "sdt", "field"),
                new JsonArray("text", "bold", "italic", "underline", "strike", "font", "size",
                    "color", "highlight", "alignment", "style", "indent", "spacing",
                    "lineSpacing", "shd", "caps", "smallcaps", "link")
            ),
            "xlsx" => (
                new JsonArray("sheet", "row", "cell", "col", "run", "shape", "chart",
                    "picture", "comment", "namedrange", "table", "validation",
                    "pivottable", "autofilter", "pagebreak", "colbreak", "cf"),
                new JsonArray("value", "formula", "bold", "italic", "font.name", "font.size",
                    "font.color", "fill", "border.all", "alignment.horizontal",
                    "alignment.vertical", "numfmt", "link", "type", "width", "height")
            ),
            _ => (new JsonArray(), new JsonArray())
        };

        obj["elements"] = elements;
        obj["commonProperties"] = commonProps;
        return obj.ToJsonString(JsonOptions);
    }

    // ==================== Element detail ====================

    private static string ElementDetail(string format, string element)
    {
        var props = GetElementProperties(format, element);
        if (props == null)
            return MakeError($"Unknown element '{element}' for format '{format}'. Call schema with only format to see available elements.");

        var obj = new JsonObject
        {
            ["format"] = format,
            ["element"] = element,
            ["properties"] = props
        };
        return obj.ToJsonString(JsonOptions);
    }

    private static JsonObject? GetElementProperties(string format, string element)
    {
        return format switch
        {
            "pptx" => GetPptxProperties(element),
            "docx" => GetDocxProperties(element),
            "xlsx" => GetXlsxProperties(element),
            _ => null
        };
    }

    // ==================== PPTX properties ====================

    private static JsonObject? GetPptxProperties(string element) => element switch
    {
        "slide" => BuildProps(
            P("layout", "string", "Slide layout name or type", example: "Blank"),
            P("background", "string", "Background: solid hex, gradient (C1-C2[-angle]), or image (image:/path)", format: "#RRGGBB", example: "#FFFFFF"),
            P("transition", "string", "Slide transition effect", values: ["fade", "push", "wipe", "split", "reveal", "random", "cover", "uncover", "zoom", "morph", "none"]),
            P("notes", "string", "Speaker notes text"),
            P("advanceTime", "number", "Auto-advance time in ms", example: "3000"),
            P("advanceClick", "boolean", "Advance on click (default true)")
        ),
        "shape" or "textbox" => BuildProps(
            P("text", "string", "Shape text content (supports \\n for line breaks)"),
            P("bold", "boolean", "Bold text", values: ["true", "false"]),
            P("italic", "boolean", "Italic text", values: ["true", "false"]),
            P("underline", "string", "Underline style", values: ["true", "single", "double", "heavy", "dotted", "dash", "wavy", "false"]),
            P("strike", "boolean", "Strikethrough text", values: ["true", "false"]),
            P("superscript", "boolean", "Superscript text"),
            P("subscript", "boolean", "Subscript text"),
            P("color", "string", "Text color (hex or theme name)", format: "#RRGGBB", example: "FF0000"),
            P("font", "string", "Font family name", example: "Arial"),
            P("size", "string", "Font size with unit", format: "number + unit", example: "24pt"),
            P("align", "string", "Text horizontal alignment", values: ["left", "center", "right", "justify"]),
            P("valign", "string", "Text vertical alignment", values: ["top", "center", "bottom"]),
            P("fill", "string", "Shape fill color (hex) or 'none'", format: "#RRGGBB", example: "4472C4"),
            P("gradient", "string", "Gradient fill: linear C1-C2[-angle] or radial:C1-C2[-focus]", example: "FF0000-0000FF-90"),
            P("image", "string", "Shape image fill (file path)"),
            P("line", "string", "Border color (hex) or 'none'", format: "#RRGGBB", example: "FF0000"),
            P("lineWidth", "string", "Border width (EMU or cm/pt)", example: "2pt"),
            P("lineDash", "string", "Border dash style", values: ["solid", "dot", "dash", "dashdot", "longdash"]),
            P("lineOpacity", "string", "Border opacity 0.0-1.0"),
            P("preset", "string", "Shape geometry preset", values: ["rect", "roundRect", "ellipse", "triangle", "diamond", "rightArrow", "chevron", "star5", "heart", "cloud"]),
            P("x", "string", "Horizontal position", format: "number + unit", units: ["cm","in","pt","emu"], example: "2cm"),
            P("y", "string", "Vertical position", format: "number + unit", units: ["cm","in","pt","emu"], example: "5cm"),
            P("width", "string", "Shape width", format: "number + unit", example: "10cm"),
            P("height", "string", "Shape height", format: "number + unit", example: "5cm"),
            P("rotation", "number", "Rotation in degrees", example: "45"),
            P("opacity", "string", "Fill opacity 0.0-1.0", example: "0.5"),
            P("margin", "string", "Text padding inside shape (e.g. 0.5cm or left,top,right,bottom)", example: "0.5cm"),
            P("lineSpacing", "string", "Line spacing", example: "1.5x"),
            P("spaceBefore", "string", "Space before paragraph", example: "12pt"),
            P("spaceAfter", "string", "Space after paragraph", example: "6pt"),
            P("list", "string", "List style", values: ["bullet", "numbered", "alpha", "roman", "none"]),
            P("indent", "string", "Paragraph first-line indent (EMU or cm/pt)"),
            P("marginLeft", "string", "Paragraph left margin (EMU or cm/pt)"),
            P("spacing", "string", "Character spacing in points", example: "2"),
            P("textFill", "string", "Text gradient fill (same format as gradient)"),
            P("textWarp", "string", "WordArt text effect or 'none'", example: "textWave1"),
            P("autoFit", "string", "Text auto-fit mode", values: ["true", "shape", "false"]),
            P("flipH", "boolean", "Horizontal flip"),
            P("flipV", "boolean", "Vertical flip"),
            P("shadow", "string", "Shadow: COLOR-BLUR-ANGLE-DIST-OPACITY or 'none'", example: "FFFFFF-6-135-4-60"),
            P("glow", "string", "Glow: COLOR-RADIUS-OPACITY or 'none'", example: "FF0000-8-75"),
            P("reflection", "string", "Reflection effect", values: ["tight", "half", "full", "none"]),
            P("softedge", "string", "Soft edge radius in pt or 'none'", example: "5"),
            P("animation", "string", "Animation: EFFECT[-CLASS][-DIRECTION][-DURATION][-TRIGGER]", example: "fade-entrance-500"),
            P("zorder", "string", "Z-order: front, back, forward, backward, or absolute position"),
            P("link", "string", "Hyperlink URL", example: "https://example.com"),
            P("alt", "string", "Alt text for accessibility")
        ),
        "picture" => BuildProps(
            P("path", "string", "Image file path (for add or to replace image)"),
            P("alt", "string", "Alt text for accessibility"),
            P("x", "string", "Horizontal position", format: "number + unit", example: "2cm"),
            P("y", "string", "Vertical position", format: "number + unit", example: "5cm"),
            P("width", "string", "Picture width", format: "number + unit", example: "10cm"),
            P("height", "string", "Picture height", format: "number + unit", example: "5cm"),
            P("rotation", "number", "Rotation in degrees", example: "0"),
            P("crop", "string", "Crop percentage: left,top,right,bottom", example: "10,10,10,10"),
            P("link", "string", "Hyperlink URL"),
            P("opacity", "number", "Opacity 0-100", range: "0-100")
        ),
        "table" => BuildProps(
            P("x", "string", "Table horizontal position", format: "number + unit", example: "2cm"),
            P("y", "string", "Table vertical position", format: "number + unit", example: "5cm"),
            P("width", "string", "Table width", format: "number + unit", example: "20cm"),
            P("height", "string", "Table height", format: "number + unit", example: "10cm"),
            P("rows", "number", "Number of rows (add only)", example: "3"),
            P("cols", "number", "Number of columns (add only)", example: "4"),
            P("tableStyle", "string", "Built-in table style", values: ["medium1", "medium2", "medium3", "medium4", "light1", "light2", "light3", "dark1", "dark2", "none"]),
            P("name", "string", "Table name")
        ),
        "row" => BuildProps(
            P("height", "string", "Row height", format: "number + unit", example: "1cm")
        ),
        "cell" => BuildProps(
            P("text", "string", "Cell text content"),
            P("bold", "boolean", "Bold text", values: ["true", "false"]),
            P("italic", "boolean", "Italic text"),
            P("color", "string", "Text color", format: "#RRGGBB", example: "#000000"),
            P("font", "string", "Font family", example: "Arial"),
            P("size", "string", "Font size", format: "number + unit", example: "12pt"),
            P("fill", "string", "Cell fill color", format: "#RRGGBB", example: "#E2EFDA"),
            P("valign", "string", "Vertical alignment", values: ["top", "center", "bottom"]),
            P("align", "string", "Horizontal alignment", values: ["left", "center", "right"]),
            P("border.all", "string", "All borders", values: ["thin", "medium", "thick", "none"]),
            P("border.color", "string", "Border color", format: "#RRGGBB"),
            P("gridspan", "number", "Horizontal span (column count, alias: colspan)"),
            P("rowspan", "number", "Vertical span (row count)"),
            P("hmerge", "boolean", "Continuation cell in horizontal merge"),
            P("vmerge", "boolean", "Continuation cell in vertical merge")
        ),
        "paragraph" => BuildProps(
            P("text", "string", "Paragraph text content"),
            P("alignment", "string", "Text alignment", values: ["left", "center", "right", "justify"]),
            P("spaceBefore", "string", "Space before paragraph", example: "12pt"),
            P("spaceAfter", "string", "Space after paragraph", example: "6pt"),
            P("lineSpacing", "string", "Line spacing", example: "1.5x"),
            P("level", "number", "Indentation level", example: "0")
        ),
        "run" => BuildProps(
            P("text", "string", "Run text content"),
            P("bold", "boolean", "Bold", values: ["true", "false"]),
            P("italic", "boolean", "Italic", values: ["true", "false"]),
            P("underline", "boolean", "Underline"),
            P("strike", "boolean", "Strikethrough"),
            P("superscript", "boolean", "Superscript"),
            P("subscript", "boolean", "Subscript"),
            P("color", "string", "Text color", format: "#RRGGBB", example: "#FF0000"),
            P("font", "string", "Font family", example: "Arial"),
            P("size", "string", "Font size", format: "number + unit", example: "24pt")
        ),
        "chart" => BuildProps(
            P("chartType", "string", "Chart type", values: ["bar", "column", "line", "pie", "scatter", "area", "doughnut", "radar"]),
            P("title", "string", "Chart title"),
            P("data", "string", "Chart data as JSON"),
            P("categories", "string", "Category labels as JSON array"),
            P("legend", "string", "Legend position", values: ["top", "bottom", "left", "right", "none"]),
            P("x", "string", "Chart horizontal position", format: "number + unit"),
            P("y", "string", "Chart vertical position", format: "number + unit"),
            P("width", "string", "Chart width", format: "number + unit"),
            P("height", "string", "Chart height", format: "number + unit"),
            P("colors", "string", "Series colors, comma-separated #RRGGBB")
        ),
        "group" or "connector" or "animation" or "video" or "equation" or "notes" or "zoom" => BuildProps(
            P("x", "string", "Horizontal position", format: "number + unit", example: "2cm"),
            P("y", "string", "Vertical position", format: "number + unit", example: "5cm"),
            P("width", "string", "Width", format: "number + unit", example: "10cm"),
            P("height", "string", "Height", format: "number + unit", example: "5cm")
        ),
        "placeholder" => BuildProps(
            P("text", "string", "Placeholder text content (supports \\n)"),
            P("bold", "boolean", "Bold text"),
            P("italic", "boolean", "Italic text"),
            P("color", "string", "Text color", format: "#RRGGBB", example: "FF0000"),
            P("font", "string", "Font family", example: "Arial"),
            P("size", "string", "Font size", format: "number + unit", example: "24pt"),
            P("align", "string", "Text alignment", values: ["left", "center", "right", "justify"]),
            P("fill", "string", "Shape fill color or 'none'"),
            P("x", "string", "Horizontal position", format: "number + unit"),
            P("y", "string", "Vertical position", format: "number + unit"),
            P("width", "string", "Width", format: "number + unit"),
            P("height", "string", "Height", format: "number + unit")
        ),
        _ => null
    };

    // ==================== DOCX properties ====================

    private static JsonObject? GetDocxProperties(string element) => element switch
    {
        "paragraph" or "p" => BuildProps(
            P("text", "string", "Paragraph text (sets single run)"),
            P("style", "string", "Paragraph style name", example: "Heading1"),
            P("alignment", "string", "Text alignment", values: ["left", "center", "right", "justify"]),
            P("firstLineIndent", "string", "First line indent in twips or with unit", example: "720"),
            P("leftIndent", "string", "Left indent in twips or with unit"),
            P("rightIndent", "string", "Right indent in twips or with unit"),
            P("hangingIndent", "string", "Hanging indent in twips or with unit"),
            P("spaceBefore", "string", "Space before", format: "unit-qualified", example: "12pt"),
            P("spaceAfter", "string", "Space after", format: "unit-qualified", example: "6pt"),
            P("lineSpacing", "string", "Line spacing", format: "multiplier, percent, or fixed", example: "1.5x"),
            P("shd", "string", "Shading", format: "fill or pattern;fill or pattern;fill;color"),
            P("listStyle", "string", "List style", values: ["bullet", "numbered", "none"]),
            P("start", "number", "Numbering start value"),
            P("numId", "number", "Numbering definition ID"),
            P("numLevel", "number", "Numbering level (ilvl)"),
            P("keepNext", "boolean", "Keep with next paragraph"),
            P("keepLines", "boolean", "Keep lines together"),
            P("pageBreakBefore", "boolean", "Page break before"),
            P("widowControl", "boolean", "Widow/orphan control")
        ),
        "run" or "r" => BuildProps(
            P("text", "string", "Run text content"),
            P("bold", "boolean", "Bold", values: ["true", "false"]),
            P("italic", "boolean", "Italic", values: ["true", "false"]),
            P("underline", "boolean", "Underline (true, single, double)", values: ["true", "single", "double", "false"]),
            P("strike", "boolean", "Strikethrough"),
            P("superscript", "boolean", "Superscript"),
            P("subscript", "boolean", "Subscript"),
            P("font", "string", "Font family", example: "Arial"),
            P("size", "string", "Font size", format: "number + unit", example: "14pt"),
            P("color", "string", "Text color", format: "#RRGGBB", example: "#FF0000"),
            P("highlight", "string", "Highlight color", values: ["yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black", "none"]),
            P("caps", "boolean", "All caps"),
            P("smallcaps", "boolean", "Small caps"),
            P("vanish", "boolean", "Hidden text"),
            P("shd", "string", "Shading"),
            P("alt", "string", "Alt text (for inline images)"),
            P("width", "string", "Image width (for inline images)", format: "number + unit", example: "5cm"),
            P("height", "string", "Image height (for inline images)", format: "number + unit")
        ),
        "table" or "tbl" => BuildProps(
            P("width", "string", "Table width", example: "100%"),
            P("alignment", "string", "Table alignment", values: ["left", "center", "right"]),
            P("style", "string", "Table style name"),
            P("rows", "number", "Number of rows (add only)"),
            P("cols", "number", "Number of columns (add only)"),
            P("indent", "string", "Table indent (twips)"),
            P("cellSpacing", "string", "Cell spacing (twips)"),
            P("layout", "string", "Table layout", values: ["fixed", "auto"]),
            P("padding", "string", "Default cell margin (twips)"),
            P("border.all", "string", "All borders", format: "style[;size[;color[;space]]]", example: "single;4;000000"),
            P("border.top", "string", "Top border"),
            P("border.bottom", "string", "Bottom border"),
            P("border.left", "string", "Left border"),
            P("border.right", "string", "Right border"),
            P("border.insideH", "string", "Inside horizontal border"),
            P("border.insideV", "string", "Inside vertical border")
        ),
        "row" or "tr" => BuildProps(
            P("height", "string", "Row height", example: "1cm"),
            P("header", "boolean", "Repeat as header row")
        ),
        "cell" or "tc" => BuildProps(
            P("text", "string", "Cell text"),
            P("font", "string", "Font family"),
            P("size", "string", "Font size", example: "12pt"),
            P("bold", "boolean", "Bold"),
            P("italic", "boolean", "Italic"),
            P("color", "string", "Text color", format: "#RRGGBB"),
            P("shd", "string", "Cell shading/fill"),
            P("alignment", "string", "Text alignment", values: ["left", "center", "right", "justify"]),
            P("valign", "string", "Vertical alignment", values: ["top", "center", "bottom"]),
            P("width", "string", "Cell width"),
            P("vmerge", "string", "Vertical merge", values: ["restart", "continue"]),
            P("gridspan", "number", "Horizontal span (column count)"),
            P("padding", "string", "Cell padding in twips (all sides)"),
            P("padding.top", "string", "Top padding (twips)"),
            P("padding.bottom", "string", "Bottom padding (twips)"),
            P("padding.left", "string", "Left padding (twips)"),
            P("padding.right", "string", "Right padding (twips)"),
            P("textDirection", "string", "Text direction", values: ["btlr", "tbrl", "lrtb", "horizontal", "vertical", "vertical-rl"]),
            P("nowrap", "boolean", "No text wrap"),
            P("border.all", "string", "All borders", format: "style[;size[;color[;space]]]", example: "single;4;000000"),
            P("border.top", "string", "Top border"),
            P("border.bottom", "string", "Bottom border"),
            P("border.left", "string", "Left border"),
            P("border.right", "string", "Right border")
        ),
        "picture" => BuildProps(
            P("src", "string", "Image file path"),
            P("alt", "string", "Alt text for accessibility"),
            P("width", "string", "Picture width", format: "number + unit", example: "10cm"),
            P("height", "string", "Picture height", format: "number + unit"),
            P("title", "string", "Picture title")
        ),
        "section" => BuildProps(
            P("orientation", "string", "Page orientation", values: ["portrait", "landscape"]),
            P("pageWidth", "string", "Page width", example: "21cm"),
            P("pageHeight", "string", "Page height", example: "29.7cm"),
            P("marginTop", "string", "Top margin"),
            P("marginBottom", "string", "Bottom margin"),
            P("marginLeft", "string", "Left margin"),
            P("marginRight", "string", "Right margin")
        ),
        "header" or "footer" => BuildProps(
            P("text", "string", "Header/footer text"),
            P("font", "string", "Font family"),
            P("size", "string", "Font size", example: "10pt"),
            P("bold", "boolean", "Bold"),
            P("italic", "boolean", "Italic"),
            P("alignment", "string", "Alignment", values: ["left", "center", "right"])
        ),
        "hyperlink" => BuildProps(
            P("text", "string", "Display text"),
            P("link", "string", "URL or bookmark reference", example: "https://example.com"),
            P("tooltip", "string", "Hover tooltip text")
        ),
        "bookmark" => BuildProps(
            P("name", "string", "Bookmark name"),
            P("text", "string", "Bookmarked text content")
        ),
        "style" => BuildProps(
            P("name", "string", "Style name", example: "MyStyle"),
            P("id", "string", "Style ID"),
            P("type", "string", "Style type", values: ["paragraph", "character", "table"]),
            P("basedon", "string", "Base style ID"),
            P("next", "string", "Next style ID"),
            P("font", "string", "Font family", example: "Arial"),
            P("size", "string", "Font size", example: "14pt"),
            P("bold", "boolean", "Bold"),
            P("italic", "boolean", "Italic"),
            P("color", "string", "Text color", format: "#RRGGBB"),
            P("alignment", "string", "Paragraph alignment", values: ["left", "center", "right", "justify"]),
            P("spacebefore", "string", "Space before (unit-qualified)", example: "12pt"),
            P("spaceafter", "string", "Space after (unit-qualified)", example: "6pt")
        ),
        "chart" => BuildProps(
            P("chartType", "string", "Chart type", values: ["column", "bar", "line", "pie", "doughnut", "area", "scatter", "bubble", "radar", "stock", "combo"]),
            P("title", "string", "Chart title text (or 'none' to remove)"),
            P("legend", "string", "Legend position", values: ["top", "bottom", "left", "right", "none"]),
            P("categories", "string", "Category labels (comma-separated)"),
            P("data", "string", "Series data: 'S1:1,2;S2:3,4'"),
            P("colors", "string", "Series colors (comma-separated hex)"),
            P("dataLabels", "string", "Data labels", values: ["value", "category", "series", "percent", "all", "none"]),
            P("width", "string", "Chart width", format: "number + unit", example: "15cm"),
            P("height", "string", "Chart height", format: "number + unit", example: "10cm")
        ),
        "footnote" => BuildProps(
            P("text", "string", "Footnote text content")
        ),
        "endnote" => BuildProps(
            P("text", "string", "Endnote text content")
        ),
        "field" => BuildProps(
            P("instruction", "string", "Field code", example: " PAGE "),
            P("text", "string", "Cached display text (alias: result)"),
            P("dirty", "boolean", "Force recalculation on open"),
            P("font", "string", "Font family"),
            P("size", "string", "Font size"),
            P("bold", "boolean", "Bold"),
            P("color", "string", "Text color")
        ),
        "sdt" or "contentcontrol" => BuildProps(
            P("sdtType", "string", "Content control type", values: ["text", "richtext", "dropdown", "combobox", "date"]),
            P("alias", "string", "Display name"),
            P("tag", "string", "Identifier tag"),
            P("lock", "string", "Lock mode", values: ["unlocked", "content", "sdt", "both"]),
            P("text", "string", "Content text"),
            P("items", "string", "Comma-separated items (for dropdown/combobox)"),
            P("format", "string", "Date format (for date type)", example: "yyyy-MM-dd")
        ),
        "watermark" => BuildProps(
            P("text", "string", "Watermark text", example: "DRAFT"),
            P("color", "string", "Watermark color", example: "silver"),
            P("font", "string", "Font family", example: "Calibri"),
            P("opacity", "string", "Opacity", example: ".5"),
            P("rotation", "number", "Rotation in degrees", example: "315")
        ),
        "equation" or "comment" or "toc" or "pagebreak" => BuildProps(
            P("text", "string", "Text content"),
            P("name", "string", "Element name or identifier")
        ),
        _ => null
    };

    // ==================== XLSX properties ====================

    private static JsonObject? GetXlsxProperties(string element) => element switch
    {
        "cell" => BuildProps(
            P("value", "string", "Cell value"),
            P("formula", "string", "Cell formula (without leading =)", example: "SUM(A1:A10)"),
            P("arrayformula", "string", "Array formula"),
            P("type", "string", "Value type", values: ["string", "number", "boolean"]),
            P("clear", "boolean", "Clear cell contents"),
            P("link", "string", "Hyperlink URL"),
            P("bold", "boolean", "Bold text"),
            P("italic", "boolean", "Italic text"),
            P("strike", "boolean", "Strikethrough"),
            P("underline", "string", "Underline style", values: ["true", "single", "double"]),
            P("superscript", "boolean", "Superscript"),
            P("subscript", "boolean", "Subscript"),
            P("font.color", "string", "Font color", format: "#RRGGBB", example: "#FF0000"),
            P("font.size", "string", "Font size", format: "number + unit", example: "14pt"),
            P("font.name", "string", "Font family", example: "Calibri"),
            P("fill", "string", "Cell fill color", format: "#RRGGBB", example: "#4472C4"),
            P("border.all", "string", "All borders", values: ["thin", "medium", "thick", "none"]),
            P("border.left", "string", "Left border"),
            P("border.right", "string", "Right border"),
            P("border.top", "string", "Top border"),
            P("border.bottom", "string", "Bottom border"),
            P("border.color", "string", "Border color", format: "#RRGGBB"),
            P("alignment.horizontal", "string", "Horizontal alignment", values: ["left", "center", "right"]),
            P("alignment.vertical", "string", "Vertical alignment", values: ["top", "center", "bottom"]),
            P("alignment.wrapText", "boolean", "Wrap text in cell"),
            P("numfmt", "string", "Number format", example: "#,##0.00"),
            P("rotation", "number", "Text rotation 0-180"),
            P("indent", "number", "Indentation level"),
            P("shrinktofit", "boolean", "Shrink to fit"),
            P("locked", "boolean", "Cell locked (sheet protection)"),
            P("formulahidden", "boolean", "Hide formula (sheet protection)")
        ),
        "sheet" => BuildProps(
            P("name", "string", "Sheet name"),
            P("freeze", "string", "Freeze panes at cell reference", example: "A2"),
            P("zoom", "number", "Zoom level 75-200", range: "75-200", example: "100"),
            P("tabcolor", "string", "Sheet tab color", format: "#RRGGBB or none", example: "#FF0000"),
            P("autofilter", "string", "AutoFilter range or none", example: "A1:F100"),
            P("merge", "string", "Merge cell range", example: "A1:D1"),
            P("protect", "boolean", "Enable sheet protection"),
            P("password", "string", "Sheet protection password"),
            P("printarea", "string", "Print area range", example: "$A$1:$D$10"),
            P("orientation", "string", "Print orientation", values: ["landscape", "portrait"]),
            P("papersize", "number", "Paper size (1=Letter, 9=A4)", example: "9"),
            P("fittopage", "string", "Fit to pages (WxH or true)", example: "1x2"),
            P("header", "string", "Page header with codes", example: "&CPage &P"),
            P("footer", "string", "Page footer with codes", example: "&LConfidential"),
            P("sort", "string", "Sort specification or none", example: "A:asc,B:desc")
        ),
        "row" => BuildProps(
            P("height", "number", "Row height in points", example: "20"),
            P("hidden", "boolean", "Hide row"),
            P("outline", "number", "Outline/group level (0-7)"),
            P("collapsed", "boolean", "Collapse outline group")
        ),
        "col" => BuildProps(
            P("width", "number", "Column width in characters", example: "15"),
            P("hidden", "boolean", "Hide column"),
            P("outline", "number", "Outline/group level (0-7)"),
            P("collapsed", "boolean", "Collapse outline group")
        ),
        "run" => BuildProps(
            P("text", "string", "Run text content"),
            P("bold", "boolean", "Bold"),
            P("italic", "boolean", "Italic"),
            P("strike", "boolean", "Strikethrough"),
            P("underline", "boolean", "Underline"),
            P("superscript", "boolean", "Superscript"),
            P("subscript", "boolean", "Subscript"),
            P("size", "string", "Font size", example: "14pt"),
            P("color", "string", "Font color", format: "#RRGGBB", example: "#FF0000"),
            P("font", "string", "Font family", example: "Calibri")
        ),
        "chart" => BuildProps(
            P("chartType", "string", "Chart type", values: ["bar", "column", "line", "pie", "scatter", "area", "doughnut", "radar", "bubble", "stock", "combo", "column3d", "bar3d"]),
            P("title", "string", "Chart title (or 'none' to remove)"),
            P("title.font", "string", "Title font typeface"),
            P("title.size", "string", "Title font size in pt"),
            P("title.color", "string", "Title font color (hex)"),
            P("title.bold", "boolean", "Title bold"),
            P("legend", "string", "Legend position", values: ["top", "bottom", "left", "right", "none"]),
            P("legendFont", "string", "Legend font: size:color:fontname", example: "9:8B949E:Helvetica Neue"),
            P("axisFont", "string", "Axis label font: size:color:fontname"),
            P("data", "string", "Chart data range or series", example: "S1:1,2;S2:3,4"),
            P("categories", "string", "Category labels (comma-separated)"),
            P("colors", "string", "Series colors (comma-separated hex)"),
            P("dataLabels", "string", "Data labels", values: ["value", "category", "series", "percent", "all", "none"]),
            P("labelPos", "string", "Data label position", values: ["center", "insideEnd", "insideBase", "outsideEnd", "top", "bottom", "left", "right"]),
            P("axisTitle", "string", "Value axis title"),
            P("catTitle", "string", "Category axis title"),
            P("axisMin", "string", "Value axis minimum"),
            P("axisMax", "string", "Value axis maximum"),
            P("gridlines", "string", "Major gridlines: true/none or color:widthPt:dash"),
            P("plotFill", "string", "Plot area fill (hex or gradient C1-C2[:angle])"),
            P("chartFill", "string", "Chart area fill"),
            P("lineWidth", "string", "Series line width in pt", example: "2.5"),
            P("style", "number", "Chart style ID (1-48)"),
            P("width", "number", "Chart width in pixels"),
            P("height", "number", "Chart height in pixels"),
            P("x", "number", "Chart position (col/row offset)"),
            P("y", "number", "Chart position (col/row offset)")
        ),
        "cf" => BuildProps(
            P("type", "string", "CF type", values: ["databar", "colorscale", "iconset", "formula", "topn", "aboveaverage", "duplicatevalues", "uniquevalues", "containstext", "dateoccurring"]),
            P("sqref", "string", "Cell range for the rule", example: "A1:A100"),
            P("range", "string", "Cell range (alias for sqref)"),
            P("color", "string", "Font color for the rule", format: "#RRGGBB"),
            P("fill", "string", "Fill color for the rule", format: "#RRGGBB"),
            P("bold", "boolean", "Bold text"),
            P("italic", "boolean", "Italic text"),
            P("strike", "boolean", "Strikethrough"),
            P("underline", "boolean", "Underline"),
            P("border", "string", "Border style", values: ["thin", "medium"]),
            P("numfmt", "string", "Number format"),
            P("rank", "number", "Top/bottom N rank (topn type)"),
            P("bottom", "boolean", "Bottom N instead of top (topn type)"),
            P("percent", "boolean", "Percent instead of count (topn type)"),
            P("below", "boolean", "Below average (aboveaverage type)"),
            P("text", "string", "Text to match (containstext type)"),
            P("period", "string", "Date period (dateoccurring type)", values: ["today", "yesterday", "tomorrow", "last7days", "thisweek", "lastweek", "thismonth", "lastmonth"])
        ),
        "shape" => BuildProps(
            P("text", "string", "Shape text content (supports \\n)"),
            P("name", "string", "Shape name"),
            P("font", "string", "Font typeface"),
            P("size", "string", "Font size in points"),
            P("bold", "boolean", "Bold"),
            P("italic", "boolean", "Italic"),
            P("color", "string", "Text color (hex)", format: "#RRGGBB"),
            P("fill", "string", "Shape fill color (hex) or 'none'"),
            P("line", "string", "Line color (hex) or 'none'"),
            P("align", "string", "Text alignment", values: ["left", "center", "right"]),
            P("preset", "string", "Shape preset (roundRect, ellipse, etc.)"),
            P("shadow", "string", "Shadow: COLOR-BLUR-ANGLE-DIST-OPACITY or 'none'"),
            P("glow", "string", "Glow: COLOR-RADIUS-OPACITY or 'none'"),
            P("reflection", "string", "Reflection", values: ["tight", "half", "full", "none"]),
            P("softEdge", "string", "Soft edge radius in pt or 'none'"),
            P("x", "number", "Horizontal position (col/row offset)"),
            P("y", "number", "Vertical position (col/row offset)"),
            P("width", "number", "Width (col/row span)"),
            P("height", "number", "Height (col/row span)")
        ),
        "picture" => BuildProps(
            P("path", "string", "Image file path (for add)"),
            P("alt", "string", "Alt text"),
            P("x", "number", "Horizontal position (col/row offset)"),
            P("y", "number", "Vertical position (col/row offset)"),
            P("width", "number", "Width (col/row span)"),
            P("height", "number", "Height (col/row span)"),
            P("rotation", "number", "Rotation angle in degrees"),
            P("shadow", "string", "Shadow: COLOR-BLUR-DIST-DIR or 'none'"),
            P("glow", "string", "Glow: COLOR-RADIUS or 'none'"),
            P("reflection", "string", "Reflection", values: ["tight", "half", "full", "none"]),
            P("softEdge", "string", "Soft edge radius in pt or 'none'")
        ),
        "comment" => BuildProps(
            P("ref", "string", "Cell reference (for add)", example: "A1"),
            P("text", "string", "Comment text"),
            P("author", "string", "Comment author")
        ),
        "namedrange" => BuildProps(
            P("name", "string", "Range name"),
            P("ref", "string", "Range reference", example: "Sheet1!$A$1:$D$10"),
            P("scope", "string", "Scope (sheet name, omit for workbook)"),
            P("comment", "string", "Range comment")
        ),
        "table" => BuildProps(
            P("ref", "string", "Table range (for add)", example: "A1:D10"),
            P("name", "string", "Table name"),
            P("displayName", "string", "Table display name"),
            P("style", "string", "Table style", example: "TableStyleMedium2"),
            P("headerRow", "boolean", "Include header row"),
            P("totalRow", "boolean", "Include total row"),
            P("columns", "string", "Column names (comma-separated)")
        ),
        "validation" => BuildProps(
            P("sqref", "string", "Cell range", example: "A1:A100"),
            P("type", "string", "Validation type", values: ["list", "whole", "decimal", "date", "time", "textLength", "custom"]),
            P("operator", "string", "Comparison operator", values: ["between", "equal", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual", "notBetween", "notEqual"]),
            P("formula1", "string", "First formula/value", example: "Yes,No,Maybe"),
            P("formula2", "string", "Second formula/value (for between)"),
            P("allowBlank", "boolean", "Allow blank cells"),
            P("showError", "boolean", "Show error alert"),
            P("errorTitle", "string", "Error dialog title"),
            P("error", "string", "Error message"),
            P("showInput", "boolean", "Show input prompt"),
            P("promptTitle", "string", "Input prompt title"),
            P("prompt", "string", "Input prompt message")
        ),
        "pivottable" => BuildProps(
            P("source", "string", "Data source range", example: "Sheet1!A1:D100"),
            P("position", "string", "Anchor cell for pivot table", example: "F1"),
            P("rows", "string", "Row fields (comma-separated header names)"),
            P("cols", "string", "Column fields"),
            P("values", "string", "Data fields with aggregation", example: "Sales:sum,Qty:count"),
            P("filters", "string", "Page/filter fields"),
            P("name", "string", "Pivot table name"),
            P("style", "string", "Pivot table style", example: "PivotStyleLight16")
        ),
        "autofilter" or "pagebreak" or "colbreak" => BuildProps(
            P("name", "string", "Element name or identifier"),
            P("value", "string", "Element value or content")
        ),
        _ => null
    };

    // ==================== Property builder helpers ====================

    private record PropDef(string Name, string Type, string Description,
        string? Format = null, string? Example = null, string? Range = null,
        string[]? Values = null, string[]? Units = null);

    private static PropDef P(string name, string type, string description,
        string? format = null, string? example = null, string? range = null,
        string[]? values = null, string[]? units = null)
        => new(name, type, description, format, example, range, values, units);

    private static JsonObject BuildProps(params PropDef[] defs)
    {
        var obj = new JsonObject();
        foreach (var d in defs)
        {
            var prop = new JsonObject
            {
                ["type"] = d.Type,
                ["description"] = d.Description
            };
            if (d.Format != null) prop["format"] = d.Format;
            if (d.Example != null) prop["example"] = d.Example;
            if (d.Range != null) prop["range"] = d.Range;
            if (d.Values != null)
            {
                var arr = new JsonArray();
                foreach (var v in d.Values) arr.Add((JsonNode?)JsonValue.Create(v));
                prop["values"] = arr;
            }
            if (d.Units != null)
            {
                var arr = new JsonArray();
                foreach (var u in d.Units) arr.Add((JsonNode?)JsonValue.Create(u));
                prop["units"] = arr;
            }
            obj[d.Name] = prop;
        }
        return obj;
    }
}
