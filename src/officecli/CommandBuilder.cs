// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;

namespace OfficeCli;

static partial class CommandBuilder
{
    public static RootCommand BuildRootCommand()
    {
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON (AI-friendly)" };
        var forceOption = new Option<bool>("--force") { Description = "Force write even if document is protected" };

        var rootCommand = new RootCommand("""
            officecli: AI-friendly CLI for Office documents (.docx, .xlsx, .pptx)

            Help navigation (start from the deepest level you know):
              officecli pptx set              All settable elements and their properties
              officecli pptx set shape        Shape properties in detail
              officecli pptx set shape.fill   Specific property format and examples

            Replace 'pptx' with 'docx' or 'xlsx'. Commands: view, get, query, set, add, raw.
            """);
        rootCommand.Add(jsonOption);

        // Session commands
        rootCommand.Add(CreateOpenCommand(jsonOption));
        rootCommand.Add(CreateCloseCommand(jsonOption));
        rootCommand.Add(CreateWatchCommand());
        rootCommand.Add(CreateUnwatchCommand());
        rootCommand.Add(CreateResidentServeCommand());

        // Read commands
        rootCommand.Add(CreateViewCommand(jsonOption));
        rootCommand.Add(CreateGetCommand(jsonOption));
        rootCommand.Add(CreateQueryCommand(jsonOption));

        // Write commands
        rootCommand.Add(CreateSetCommand(jsonOption, forceOption));
        rootCommand.Add(CreateAddCommand(jsonOption, forceOption));
        rootCommand.Add(CreateRemoveCommand(jsonOption));
        rootCommand.Add(CreateMoveCommand(jsonOption));

        // Raw XML commands
        rootCommand.Add(CreateRawCommand(jsonOption));
        rootCommand.Add(CreateRawSetCommand(jsonOption));
        rootCommand.Add(CreateAddPartCommand(jsonOption));

        // Validation commands
        rootCommand.Add(CreateValidateCommand(jsonOption));
        rootCommand.Add(CreateCheckCommand(jsonOption));

        // Schema command (AI agent property discovery)
        rootCommand.Add(CreateSchemaCommand());

        // Batch & utility commands
        rootCommand.Add(CreateBatchCommand(jsonOption));
        rootCommand.Add(CreateImportCommand(jsonOption));
        rootCommand.Add(CreateCreateCommand(jsonOption));
        rootCommand.Add(CreateMergeCommand(jsonOption));

        HelpCommands.Register(rootCommand);

        return rootCommand;
    }
}
