# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is **ExcelToolDotNet**, a .NET 9.0 C# command-line tool that converts Excel spreadsheets (game configuration tables) into multiple output formats for use in game development. It reads Excel files (.xls/.xlsx) using the NPOI library and converts them based on XML configuration files.

## Build Commands

- **Build Debug**: `dotnet build ExcelTool.csproj -c Debug`
- **Build Release**: `dotnet build ExcelTool.csproj -c Release`
- **Clean**: `dotnet clean ExcelTool.csproj -c Debug`
- **Publish Windows x64 (trimmed)**: `dotnet publish ExcelTool.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:PublishTrimmed=true -p:TrimMode=partial`
- **Publish macOS ARM64**: `dotnet publish ExcelTool.csproj -c Release -r osx-arm64 --self-contained -p:PublishSingleFile=true`

## Run Commands

The tool is run from the command line with optional flags:

```bash
# Basic run
dotnet run --project ExcelTool.csproj

# With flags
dotnet run -- -fast -nolog -use_xls_cache

# Generate i18n translation entries
dotnet run -- -extra_text

# Clean extra text entries
dotnet run -- -clean_extra_text

# Special channel (region-specific override)
dotnet run -- -special=cn

# Use test data
dotnet run -- -use_test_data

# Dynamic output mode
dotnet run -- -dynamic_output

# Output C# access interface
dotnet run -- -csharp
```

## Directory Structure

- `ExcelTool/` - Main C# source files
- `libxl/` - Excel reading library bindings
- `config/` - XML configuration files (table definitions, enums)
- `source/` - Input Excel .xls/.xlsx files
- `output_*/` - Various output directories created by the tool

## Architecture

### Core Classes

1. **Program.cs** - Entry point, command-line argument parsing
2. **ConvertTool.cs** - Main conversion orchestrator
3. **FieldConfig.cs** - Table and field configuration (loaded from XML)
4. **ConvertConfig.cs** - Data structures: `ExcelField`, `InputConfig`, `OutputConfig`, `EnumItem`

### Output Format Converters (partial classes)

The `ConvertTool` class is split across multiple files using C# partial classes:

- `ConvertTool_Lua.cs` - Lua table output
- `ConvertTool_BinLua.cs` - Binary Lua output for client
- `ConvertTool_CSharp.cs` - C# class output
- `ConvertTool_Cpp.cs` - C++ struct output
- `ConvertTool_Go.cs` - Go struct output
- `ConvertTool_Javascript.cs` - JSON output
- `ConvertTool_xml.cs` - XML output
- `ConvertTool_HumanReadable.cs` - Plain text human-readable output
- `ConvertTool_ScriptableObject.cs` - Unity ScriptableObject YAML output
- `ConvertTool_Bin.cs` - Binary output for server

### Support Classes

- **Assist.cs** - Type conversion helpers (C++/C#/Go types)
- **BaseHelper.cs** - File I/O utilities (WriteText, WriteBin, etc.)
- **EnumManager.cs** - Enum type handling for export
- **CustomEnum.cs** - Runtime custom enum management
- **SheetCache.cs** - Excel sheet data caching
- **SheetCacheMgr.cs** - Cache manager for Excel files
- **XlsLoader.cs** - Excel file loading via NPOI
- **Log.cs** - Logging utilities
- **GlobeError.cs** - Global error tracking
- **I18N.cs** - Internationalization/text extraction

## Key Concepts

### XML Table Configuration

Each table is defined in a `config/*.xml` file with:
- `<table name="..." desc="..." use_mode="Common|Client|Server" export_xml="1" export_csharp="1" export_golang="1" export_enum_only="0">` - Table metadata with export flags
- `<input>` - Container for source Excel file(s) - any child element with `file` attribute works (e.g., `<item file="...">`)
- `<output file="...">` - Output filename
- `<fields>` - Container for field definitions - any child element works. Field attributes:
  - `key` - Excel column header to match
  - `type` - Data type: `int`, `string`, `number`/`double`, `centimeter`, `decimeter`, `millimetre`, `ratio`
  - `name` - Output field name
  - `primary="true"` - Primary key field
  - `text="true"` - Mark for i18n translation extraction
  - `raw_string="true"` - **(string only)** Skip string table and i18n extraction entirely (no value lookup, no translation extraction)
  - `client_only="true"` - Skip server-side export
  - `export_bin="true"` - Include in binary export
  - `ref_table`, `ref_column` - Cross-table reference validation
  - `min_num`, `max_num` - Value range validation
  - And many more for value comparison validation

### Field Types

Common field types: `string`, `int`, `number`/`double`, `centimeter`, `decimeter`, `millimetre`, `ratio`

### Example Configuration

```xml
<config>
    <table name="ActivityAccumulateLoginTableData" export_xml="1" use_mode="Common">
        <input>
            <item file="活动/累计登录活动表.xls"/>
        </input>
        <output file="ActivityAccumulateLoginTableData.csv"/>
        <fields>
            <item name="ID" key="ID" type="int" primary="true"/>
            <item name="Day" key="天数" type="int"/>
            <item name="BGImage" key="背景图" type="string" raw_string="true"/>
            <item name="AwardType" key="奖励类型" type="string" text="true"/>
            <item name="AwardNum" key="奖励数量" type="int"/>
        </fields>   
    </table>
</config>
```

**Note**: The code iterates all child elements regardless of tag name (`<item>` works fine), only checking for expected attributes like `file`, `key`, `type`, `name`, etc.

### Data Flow

1. Load `config/enums.xml` for global enum definitions
2. Load each `config/*.xml` table configuration
3. Pre-check validation (cross-table references)
4. Read Excel files from `source/` directory
5. Map Excel columns to configured fields
6. Convert data to each output format
7. Write output files to `output_*/` directories

## Important Notes

- The tool expects a `config/` directory with XML table definitions and an `enums.xml` file
- The tool expects a `source/` directory with Excel .xls/.xlsx files
- Uses libxl bindings for Excel reading
- Supports xls/xlsx fallback (tries alternate extension if file not found)
- Internationalization (i18n): fields marked with `text="true"` are extracted for translation
- Supports special/regional channel overrides (looks in `{channel}/source/` first)
