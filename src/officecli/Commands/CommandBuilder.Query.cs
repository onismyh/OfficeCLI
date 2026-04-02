// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command CreateQueryCommand(Option<bool> jsonOption)
    {
        var queryFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var selectorArg = new Argument<string>("selector") { Description = "CSS-like selector (e.g. paragraph[style=Normal] > run[font!=Arial])" };

        var queryTextOpt = new Option<string?>("--text") { Description = "Filter results to elements containing this text (case-insensitive)" };

        var queryCommand = new Command("query", "Query document elements with CSS-like selectors");
        queryCommand.Add(queryFileArg);
        queryCommand.Add(selectorArg);
        queryCommand.Add(jsonOption);
        queryCommand.Add(queryTextOpt);

        queryCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(queryFileArg)!;
            var selector = result.GetValue(selectorArg)!;
            var textFilter = result.GetValue(queryTextOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "query";
                req.Json = json;
                req.Args["selector"] = selector;
                if (textFilter != null) req.Args["text"] = textFilter;
            }, json) is {} rc) return rc;

            var format = json ? OutputFormat.Json : OutputFormat.Text;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var filters = OfficeCli.Core.AttributeFilter.Parse(selector);
            var (results, warnings) = OfficeCli.Core.AttributeFilter.ApplyWithWarnings(handler.Query(selector), filters);
            if (!string.IsNullOrEmpty(textFilter))
                results = results.Where(n => n.Text != null && n.Text.Contains(textFilter, StringComparison.OrdinalIgnoreCase)).ToList();
            if (json)
            {
                var cliWarnings = warnings.Select(w => new OfficeCli.Core.CliWarning { Message = w, Code = "filter_warning" }).ToList();
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNodes(results, OutputFormat.Json),
                    cliWarnings.Count > 0 ? cliWarnings : null));
            }
            else
            {
                foreach (var w in warnings) Console.Error.WriteLine(w);
                Console.WriteLine(OutputFormatter.FormatNodes(results, OutputFormat.Text));
            }
            return 0;
        }, json); });

        return queryCommand;
    }
}
