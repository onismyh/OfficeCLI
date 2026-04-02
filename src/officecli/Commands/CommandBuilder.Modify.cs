// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command CreateRemoveCommand(Option<bool> jsonOption)
    {
        var removeFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var removePathArg = new Argument<string>("path") { Description = "DOM path of the element to remove" };

        var removeCommand = new Command("remove", "Remove an element from the document");
        removeCommand.Add(removeFileArg);
        removeCommand.Add(removePathArg);
        removeCommand.Add(jsonOption);

        removeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(removeFileArg)!;
            var path = result.GetValue(removePathArg)!;

            if (TryResident(file.FullName, req =>
            {
                req.Command = "remove";
                req.Args["path"] = path;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var oldCount = (handler as OfficeCli.Handlers.PowerPointHandler)?.GetSlideCount() ?? 0;
            var warning = handler.Remove(path);
            var message = $"Removed {path}";
            if (warning != null) message += $"\n{warning}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
            else Console.WriteLine(message);
            var slideNum = WatchMessage.ExtractSlideNum(path);
            if (slideNum > 0 && !path.Contains("/shape["))
                NotifyWatchRoot(handler, file.FullName, oldCount);
            else
                NotifyWatch(handler, file.FullName, path);
            return 0;
        }, json); });

        return removeCommand;
    }

    private static Command CreateMoveCommand(Option<bool> jsonOption)
    {
        var moveFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var movePathArg = new Argument<string>("path") { Description = "DOM path of the element to move" };
        var moveToOpt = new Option<string?>("--to") { Description = "Target parent path. If omitted, reorders within the current parent" };
        var moveIndexOpt = new Option<int?>("--index") { Description = "Insert position (0-based). If omitted, appends to end" };

        var moveCommand = new Command("move", "Move an element to a new position or parent");
        moveCommand.Add(moveFileArg);
        moveCommand.Add(movePathArg);
        moveCommand.Add(moveToOpt);
        moveCommand.Add(moveIndexOpt);
        moveCommand.Add(jsonOption);

        moveCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(moveFileArg)!;
            var path = result.GetValue(movePathArg)!;
            var to = result.GetValue(moveToOpt);
            var index = result.GetValue(moveIndexOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "move";
                req.Args["path"] = path;
                if (to != null) req.Args["to"] = to;
                if (index.HasValue) req.Args["index"] = index.Value.ToString();
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var resultPath = handler.Move(path, to, index);
            var message = $"Moved to {resultPath}";
            if (json)
            {
                var moveNode = handler.Get(resultPath);
                Console.WriteLine(OutputFormatter.WrapEnvelopeWithNode(message, moveNode));
            }
            else Console.WriteLine(message);
            NotifyWatch(handler, file.FullName, path);
            return 0;
        }, json); });

        return moveCommand;
    }
}
