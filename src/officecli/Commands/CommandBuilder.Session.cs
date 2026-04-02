// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Diagnostics;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command CreateOpenCommand(Option<bool> jsonOption)
    {
        var openFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var openCommand = new Command("open", "Start a resident process to keep the document in memory for faster subsequent commands");
        openCommand.Add(openFileArg);
        openCommand.Add(jsonOption);

        openCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(openFileArg)!;
            var filePath = file.FullName;

            // If already running, reuse the existing resident
            if (ResidentClient.TryConnect(filePath, out _))
            {
                var msg = $"Opened {file.Name} (already running, do NOT call close)";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else Console.WriteLine(msg);
                return 0;
            }

            // Fork a background process running the resident server
            var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
            if (exePath == null)
                throw new InvalidOperationException("Cannot determine executable path.");

            var startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"__resident-serve__ \"{filePath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            var process = Process.Start(startInfo);
            if (process == null)
                throw new InvalidOperationException("Failed to start resident process.");

            // Wait briefly for the server to start accepting connections
            for (int i = 0; i < 50; i++) // up to 5 seconds
            {
                Thread.Sleep(100);
                if (ResidentClient.TryConnect(filePath, out _))
                {
                    var msg = $"Opened {file.Name} (remember to call close when done)";
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                    else Console.WriteLine(msg);
                    return 0;
                }
                if (process.HasExited)
                {
                    var stderr = process.StandardError.ReadToEnd();
                    throw new InvalidOperationException($"Resident process exited. {stderr}");
                }
            }

            throw new InvalidOperationException("Resident process started but not responding.");
        }, json); });

        return openCommand;
    }

    private static Command CreateCloseCommand(Option<bool> jsonOption)
    {
        var closeFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var closeCommand = new Command("close", "Stop the resident process for the document");
        closeCommand.Add(closeFileArg);
        closeCommand.Add(jsonOption);

        closeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(closeFileArg)!;
            if (ResidentClient.SendClose(file.FullName))
            {
                var msg = $"Resident closed for {file.Name}";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else Console.WriteLine(msg);
            }
            else
            {
                throw new InvalidOperationException($"No resident running for {file.Name}");
            }
            return 0;
        }, json); });

        return closeCommand;
    }

    private static Command CreateResidentServeCommand()
    {
        var serveFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var serveCommand = new Command("__resident-serve__", "Internal: run resident server (do not call directly)");
        serveCommand.Hidden = true;
        serveCommand.Add(serveFileArg);

        serveCommand.SetAction(result =>
        {
            var file = result.GetValue(serveFileArg)!;
            using var server = new ResidentServer(file.FullName);
            server.RunAsync().GetAwaiter().GetResult();
        });

        return serveCommand;
    }
}
