// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command CreateImportCommand(Option<bool> jsonOption)
    {
        var importFileArg = new Argument<FileInfo>("file") { Description = "Target Excel file (.xlsx)" };
        var importParentPathArg = new Argument<string>("parent-path") { Description = "Sheet path (e.g. /Sheet1)" };
        var importSourceArg = new Argument<FileInfo?>("source-file") { Description = "Source CSV/TSV file to import (positional, alternative to --file)" };
        importSourceArg.DefaultValueFactory = _ => null!;
        var importSourceOpt = new Option<FileInfo?>("--file") { Description = "Source CSV/TSV file to import" };
        var importStdinOpt = new Option<bool>("--stdin") { Description = "Read CSV/TSV data from stdin" };
        var importFormatOpt = new Option<string?>("--format") { Description = "Data format: csv or tsv (default: inferred from file extension, or csv)" };
        var importHeaderOpt = new Option<bool>("--header") { Description = "First row is header: set AutoFilter and freeze pane" };
        var importStartCellOpt = new Option<string>("--start-cell") { Description = "Starting cell (default: A1)" };
        importStartCellOpt.DefaultValueFactory = _ => "A1";

        var importCommand = new Command("import", "Import CSV/TSV data into an Excel sheet");
        importCommand.Add(importFileArg);
        importCommand.Add(importParentPathArg);
        importCommand.Add(importSourceArg);
        importCommand.Add(importSourceOpt);
        importCommand.Add(importStdinOpt);
        importCommand.Add(importFormatOpt);
        importCommand.Add(importHeaderOpt);
        importCommand.Add(importStartCellOpt);
        importCommand.Add(jsonOption);

        importCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(importFileArg)!;
            var parentPath = result.GetValue(importParentPathArg)!;
            var source = result.GetValue(importSourceOpt) ?? result.GetValue(importSourceArg);
            var useStdin = result.GetValue(importStdinOpt);
            var format = result.GetValue(importFormatOpt);
            var header = result.GetValue(importHeaderOpt);
            var startCell = result.GetValue(importStartCellOpt)!;

            if (!file.Exists)
                throw new CliException($"File not found: {file.FullName}")
                {
                    Code = "file_not_found",
                    Suggestion = $"Create the file first: officecli create \"{file.FullName}\""
                };

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext != ".xlsx")
                throw new CliException("Import is only supported for .xlsx files in V1")
                {
                    Code = "unsupported_type",
                    Suggestion = "Use a .xlsx file"
                };

            // Read CSV content
            string csvContent;
            if (useStdin)
            {
                csvContent = Console.In.ReadToEnd();
            }
            else if (source != null)
            {
                if (!source.Exists)
                    throw new CliException($"Source file not found: {source.FullName}")
                    {
                        Code = "file_not_found"
                    };
                csvContent = File.ReadAllText(source.FullName, Encoding.UTF8);
            }
            else
            {
                throw new CliException("Either --file or --stdin must be specified")
                {
                    Code = "missing_argument",
                    Suggestion = "Use --file <path> to specify a CSV/TSV file, or --stdin to read from standard input"
                };
            }

            // Determine delimiter: --format flag > file extension > default csv
            char delimiter = ',';
            if (!string.IsNullOrEmpty(format))
            {
                delimiter = format.ToLowerInvariant() switch
                {
                    "tsv" => '\t',
                    "csv" => ',',
                    _ => throw new CliException($"Unknown format: {format}. Use 'csv' or 'tsv'")
                    {
                        Code = "invalid_value",
                        ValidValues = ["csv", "tsv"]
                    }
                };
            }
            else if (source != null)
            {
                var sourceExt = Path.GetExtension(source.FullName).ToLowerInvariant();
                if (sourceExt == ".tsv" || sourceExt == ".tab")
                    delimiter = '\t';
            }

            using var handler = new OfficeCli.Handlers.ExcelHandler(file.FullName, editable: true);
            var msg = handler.Import(parentPath, csvContent, delimiter, header, startCell);
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
            else
                Console.WriteLine(msg);
            return 0;
        }, json); });

        return importCommand;
    }
}
