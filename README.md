# MsgToEmlConverter [![.NET](https://github.com/djgosnell/MsgToEmlConverter/actions/workflows/dotnet.yml/badge.svg)](https://github.com/djgosnell/MsgToEmlConverter/actions/workflows/dotnet.yml)

A lightweight .NET console application that converts Microsoft Outlook .msg files to standard .eml format.

## Features

- Convert single .msg files to .eml format
- Batch convert entire directories of .msg files
- Preserves all email metadata (sender, recipients, subject, date)
- Maintains both plain text and HTML body content
- Handles attachments with proper MIME types
- Progress tracking for batch operations
- Optimized single-file executable

## Usage

### No Arguments (Default Mode)
```bash
MsgToEmlConverter
```
Converts all .msg files in the executable's directory to a `converted` subfolder.

### Single File Conversion
```bash
MsgToEmlConverter input.msg output.eml
```

### Directory Conversion
```bash
MsgToEmlConverter input_directory output_directory
```

## Requirements

- Windows x64
- .NET 9.0 (self-contained, no installation required)

## Building from Source

```bash
# Build the project
dotnet build

# Create optimized release build
dotnet build -c Release

# Publish as single-file executable
dotnet publish -c Release
```

The published executable will be located in `bin/Release/net9.0/win-x64/publish/`.

## Dependencies

- [MimeKit](https://github.com/jstedfast/MimeKit) - RFC-compliant email message handling
- [MSGReader](https://github.com/Sicos1977/MSGReader) - Microsoft .msg file parsing

## Output

The converter creates standard .eml files that can be opened in any email client that supports the format, including:
- Outlook
- Thunderbird
- Apple Mail
- Windows Mail
- Most web-based email clients