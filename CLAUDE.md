# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is **ExcelToolDotNet**, a .NET 9.0 C# command-line tool that converts Excel spreadsheets (game configuration tables) into multiple output formats for use in game development. It reads Excel files (.xls/.xlsx) using the NPOI and libxl libraries and converts them based on XML configuration files. Uses `System.CommandLine` for CLI argument parsing.

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

# Clean extra text entries (use with -extra_text)
dotnet run -- -clean_extra_text

# Special channel (region-specific override)
dotnet run -- -special=cn

# Use test data
dotnet run -- -use_test_data

# Dynamic output mode
dotnet run -- -dynamic_output

# Output C# access interface
dotnet run -- -csharp

# Use CSV translation file for server-side text
dotnet run -- -csv_translation=path/to/file.xlsx
```

## Directory Structure

- `ExcelTool/` - Main C# source files
- `libxl/` - Excel reading library bindings (native)
- `config/` - XML configuration files (table definitions, enums)
- `source/` - Input Excel .xls/.xlsx files
- `i18n/` - Translation Excel files (per-language xlsx) and `layoutText.json` (UI layout strings)
- `languageTables/` - Output UILayout translation JSON files
- `output_*/` - Various output directories created by the tool
- `output_asset/` - Unity ScriptableObject .asset files for string tables
- `output_textmeshpro_text/` - Character set file for TextMeshPro/SDF font generation

## Architecture

### Core Classes

1. **Program.cs** - Entry point, `System.CommandLine`-based CLI argument parsing
2. **ConvertTool.cs** - Main conversion orchestrator
3. **FieldConfig.cs** - Table and field configuration (loaded from XML), slot mapping
4. **ConvertConfig.cs** - Data structures: `ExcelField`, `InputConfig`, `OutputConfig`, `EnumItem`, `TableUseMode` enum

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

### Excel Reading & Data Processing (partial classes)

- `ExcelRead.cs` - Core Excel row reading logic: `ReadExcelRawLine()`, `CellDataForLua` struct, CSV translation loading, duplicate primary key handling (`ReadRawLineResult` enum), cross-table comparison validation (`CompareNumIsTrue`)
- `CsvRead.cs` - Legacy CSV-style row reading: `ReadCsvLine()`, `ReadErlCsvLine()` (server-only, skips `client_only` fields)
- `PreCheck.cs` - Pre-check phase: loads config + Excel, builds `allLoadCfgData` for cross-table reference validation

### Code Generation (partial class of FieldConfig)

- `FieldConfigAppend.cs` - Generates struct/class definitions for all target languages: C++ (`AppendCppDefine`), C# access layer (`AppendCSharpAccess`, `AppendCSharpDefine`), Go (`AppendGoDefine`, `AppendGoDefine2`, `AppendGoFunc`, `AppendGoDeclaration`, `AppendGoImpl`), Erlang (`AppendErlangDefine`, `AppendErlangImpl`), Lua meta (`AppendLuaDefine`), XML wrapper (`AppendXmlDefind`)

### Support Classes

- **Assist.cs** - Type conversion helpers (C++/C#/Go types, Lua string escaping)
- **BaseHelper.cs** - File I/O utilities (WriteText, WriteBin, etc.)
- **EnumManager.cs** - Enum type handling for export
- **CustomEnum.cs** - Runtime custom enum management
- **SheetCache.cs** - Excel sheet data caching (wraps libxl Sheet)
- **SheetCacheMgr.cs** - Cache manager for Excel files
- **XlsLoader.cs** - Excel file loading (libxl Book creation)
- **Log.cs** - Logging utilities
- **GlobeError.cs** - Global error tracking and reporting
- **I18N.cs** - Internationalization: text registration, ScriptableObject language table generation, UILayout JSON output, i18n sync (writes untranslated entries to xlsx files), SDF font character set generation
- **JsonContext.cs** - Source-generated JSON serialization context (`ExcelToolJsonContext`) for trimmed/AOT builds, plus `UILayoutEntry` data class and `JsonSerializerHelper` extension methods
- **Def.cs** - Erlang implementation template constants (`ERL_IMPL_BEGIN`, `ERL_IMPL_END`)

## Key Concepts

### XML Table Configuration

Each table is defined in a `config/*.xml` file with:
- `<table name="..." desc="..." use_mode="Common|Client|Server" export_xml="1" export_csharp="1" export_golang="1" export_erl="1" export_enum_only="0">` - Table metadata with export flags
- `<input>` - Container for source Excel file(s) - any child element with `file` attribute works (e.g., `<item file="...">`). Supports `dynamic="true"` attribute for dynamic source files.
- `<output file="...">` - Output filename
- `<fields>` - Container for field definitions - any child element works. Field attributes:
  - `key` - Excel column header to match
  - `type` - Data type: `int`, `string`, `number`/`double`, `centimeter`, `decimeter`, `millimetre`, `ratio`
  - `name` - Output field name
  - `primary="true"` - Primary key field
  - `ignore_duplicated="true"` - **(primary keys only)** Allow duplicate primary keys by skipping the duplicate row instead of erroring
  - `text="true"` - Mark for i18n translation extraction
  - `raw_string="true"` - **(string only)** Skip string table and i18n extraction entirely (no value lookup, no translation extraction)
  - `sdf_text="true"` - Register text characters for SDF font generation
  - `client_only="true"` - Skip server-side export
  - `export_bin="true"` - Include in binary export
  - `need_load="true"` - Mark field as needing to be loaded into memory
  - `enum_value="true"` - Mark as enum value field
  - `ref_table`, `ref_column` - Cross-table reference validation
  - `min_num`, `max_num` - Value range validation
  - `i_should_be_bigger_than_t`, `t_should_be_bigger_than_i`, `i_should_like_t` - Cross-table numeric comparison validation (used with `self_key`, `target_key`, `target_compare`, `ref_table`)

### Field Types

Common field types: `string`, `int`, `number`/`double`, `centimeter`, `decimeter`, `millimetre`, `ratio`

Any type not matching the above is treated as an enum type name, resolved via `EnumManager`.

### CellDataForLua and String Indexing

When reading Excel rows, string fields (without `raw_string`) are registered with `I18N.RegisterText()` which returns a string table index. The `CellDataForLua` struct wraps each cell value for the client pipeline, distinguishing between:
- `Standard` - plain value
- `StringIndex` - references into the i18n string table by index

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
3. Pre-check validation (cross-table references via `PreCheck.cs`)
4. Read Excel files from `source/` directory (with xls/xlsx fallback)
5. Map Excel columns to configured fields via `FieldConfig.LoadSlotInfo()`
6. Read rows via `ExcelRead.ReadExcelRawLine()` - produces separate client (`CellDataForLua`) and server data lines
7. Convert data to each output format via partial class converters
8. Write output files to `output_*/` directories
9. If `-extra_text`: sync untranslated entries into `i18n/*.xlsx` files; otherwise write language tables

## I18N System

- **Text registration**: `I18N.RegisterText()` collects all non-`raw_string` string values, assigns unique indices
- **String table output**: Writes Unity ScriptableObject `.asset` files (YAML format) to `output_asset/`
- **Translation input**: Reads `i18n/*.xlsx` files (language detected by filename: 简体→CN, 繁体→TW, 英文→EN, 日文→JP, 韩文→KR)
- **UILayout**: Reads `i18n/layoutText.json`, outputs per-language translations as JSON to `languageTables/`
- **i18n sync** (`-extra_text`): Appends untranslated entries to existing `i18n/*.xlsx` files using NPOI
- **SDF font**: Collects all characters from translated text, outputs to `output_textmeshpro_text/TextMeshPro.txt`
- **CSV translation** (`-csv_translation`): Per-table server-side text substitution from a separate Excel file

## Important Notes

- The tool expects a `config/` directory with XML table definitions and an `enums.xml` file
- The tool expects a `source/` directory with Excel .xls/.xlsx files
- Uses both libxl (native, for Sheet/Book) and NPOI (for i18n sync writing to xlsx)
- Supports xls/xlsx fallback (tries alternate extension if file not found)
- Internationalization (i18n): fields marked with `text="true"` are extracted for translation; `raw_string="true"` bypasses this
- Supports special/regional channel overrides (looks in `{channel}/source/` first)
- JSON serialization uses source-generated `ExcelToolJsonContext` to support trimmed/AOT builds
- `ImplicitUsings` is disabled; all `using` statements are explicit
