// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command CreateGetCommand(Option<bool> jsonOption)
    {
        var getFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var pathArg = new Argument<string>("path") { Description = "DOM path (e.g. /body/p[1])" };
        pathArg.DefaultValueFactory = _ => "/";
        var depthOpt = new Option<int>("--depth") { Description = "Depth of child nodes to include" };
        depthOpt.DefaultValueFactory = _ => 1;

        var getCommand = new Command("get", "Get a document node by path");
        getCommand.Add(getFileArg);
        getCommand.Add(pathArg);
        getCommand.Add(depthOpt);
        getCommand.Add(jsonOption);

        getCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(getFileArg)!;
            var path = result.GetValue(pathArg)!;
            var depth = result.GetValue(depthOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "get";
                req.Json = json;
                req.Args["path"] = path;
                req.Args["depth"] = depth.ToString();
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var node = handler.Get(path, depth);
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNode(node, OutputFormat.Json)));
            else
                Console.WriteLine(OutputFormatter.FormatNode(node, OutputFormat.Text));
            return 0;
        }, json); });

        return getCommand;
    }
}
