# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MsgToEmlConverter is a .NET 9.0 console application that converts Microsoft Outlook .msg files to standard .eml format. The application uses MimeKit for email message handling and MSGReader for parsing .msg files.

## Common Commands

### Build and Run
```bash
# Build the project
dotnet build

# Build in Release mode
dotnet build -c Release

# Run the application
dotnet run

# Run with arguments (single file conversion)
dotnet run -- input.msg output.eml

# Run with arguments (directory conversion)
dotnet run -- input_directory output_directory
```

### Publishing
```bash
# Publish as single-file executable (configured for win-x64)
dotnet publish -c Release

# The published executable will be in bin/Release/net9.0/win-x64/publish/
```

## Architecture

This is a single-file console application with the following structure:

- **Main method**: Handles command-line argument parsing and determines conversion mode (single file vs directory)
- **ConvertSingleFile**: Core conversion logic that reads .msg files using MSGReader and creates .eml files using MimeKit
- **ConvertDirectory**: Batch processing for converting multiple .msg files with progress tracking
- **CreateProgressBar**: Utility for displaying conversion progress

### Key Dependencies
- **MimeKit 4.13.0**: RFC-compliant email message construction and serialization
- **MSGReader 6.0.3**: Microsoft .msg file parsing
- **System.Text.Encoding.CodePages 9.0.8**: Extended encoding support

### Conversion Process
1. Parse .msg file using MSGReader's Storage.Message
2. Extract message properties (sender, recipients, subject, date)
3. Convert body content (both text and HTML)
4. Process attachments with proper MIME types
5. Build MimeMessage using MimeKit's BodyBuilder
6. Write to .eml file format

## Project Configuration

The project is configured with aggressive size optimization:
- Self-contained deployment
- Single-file publishing with compression
- Full trimming enabled
- Various runtime features disabled for smaller footprint
- Target: win-x64 runtime