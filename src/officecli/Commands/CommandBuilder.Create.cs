// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command CreateCreateCommand(Option<bool> jsonOption)
    {
        var createFileArg = new Argument<string>("file") { Description = "Output file path (.docx, .xlsx, .pptx)" };
        var createTypeOpt = new Option<string>("--type") { Description = "Document type (docx, xlsx, pptx) — optional, inferred from file extension" };
        var createCommand = new Command("create", "Create a blank Office document");
        createCommand.Aliases.Add("new");
        createCommand.Add(createFileArg);
        createCommand.Add(createTypeOpt);
        createCommand.Add(jsonOption);

        createCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(createFileArg)!;
            var type = result.GetValue(createTypeOpt);

            // If file has no extension but --type is provided, append it
            if (!string.IsNullOrEmpty(type) && string.IsNullOrEmpty(Path.GetExtension(file)))
            {
                var ext = type.StartsWith('.') ? type : "." + type;
                file += ext;
            }

            // Check if the file is held by a resident process
            var fullPath = Path.GetFullPath(file);
            if (ResidentClient.TryConnect(fullPath, out _))
            {
                throw new CliException($"{Path.GetFileName(file)} is currently opened by a resident process. Please run 'officecli close \"{file}\"' first.")
                {
                    Code = "file_locked",
                    Suggestion = $"Run: officecli close \"{file}\""
                };
            }

            OfficeCli.BlankDocCreator.Create(file);
            var fullCreatedPath = Path.GetFullPath(file);
            if (json)
            {
                var ext = Path.GetExtension(file).ToLowerInvariant();
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["message"] = $"Created: {fullCreatedPath}",
                    ["file"] = fullCreatedPath,
                    ["format"] = ext.TrimStart('.')
                };
                if (ext == ".pptx")
                {
                    envelope["totalSlides"] = 0;
                    envelope["slideWidth"] = Core.EmuConverter.FormatEmu(12192000);
                    envelope["slideHeight"] = Core.EmuConverter.FormatEmu(6858000);
                }
                else if (ext == ".xlsx")
                {
                    envelope["sheets"] = new System.Text.Json.Nodes.JsonArray("Sheet1");
                }
                Console.WriteLine(envelope.ToJsonString(OutputFormatter.PublicJsonOptions));
            }
            else
            {
                Console.WriteLine($"Created: {file}");
                if (Path.GetExtension(file).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"  totalSlides: 0");
                    Console.WriteLine($"  slideWidth: {Core.EmuConverter.FormatEmu(12192000)}");
                    Console.WriteLine($"  slideHeight: {Core.EmuConverter.FormatEmu(6858000)}");
                }
            }
            return 0;
        }, json); });

        return createCommand;
    }
}
