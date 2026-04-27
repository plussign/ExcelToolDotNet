using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelTool
{
    /// <summary>
    /// Source-generated JSON serialization context for trimmed/AOT builds.
    /// Required because reflection-based serialization is disabled in trimmed builds.
    /// </summary>
    [JsonSourceGenerationOptions(
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = true)]
    [JsonSerializable(typeof(List<string>))]
    [JsonSerializable(typeof(List<UILayoutEntry>))]
    internal partial class ExcelToolJsonContext : JsonSerializerContext
    {
    }

    /// <summary>
    /// Data class for UI layout translation entries.
    /// </summary>
    public class UILayoutEntry
    {
        [JsonPropertyName("text")]
        public string Text { get; set; }

        [JsonPropertyName("translation")]
        public string Translation { get; set; }
    }

    /// <summary>
    /// Helper extension methods for source-generated JSON serialization.
    /// </summary>
    internal static class JsonSerializerHelper
    {
        /// <summary>
        /// Serialize a list of UILayoutEntry objects using source-generated JSON.
        /// </summary>
#pragma warning disable IL2026
        public static string SerializeUILayoutEntries(List<UILayoutEntry> entries)
        {
            var context = ExcelToolJsonContext.Default;
            var options = new JsonSerializerOptions { TypeInfoResolver = context };
            return JsonSerializer.Serialize(entries, options);
        }
#pragma warning restore IL2026

        /// <summary>
        /// Deserialize a JSON string to a list of strings using source-generated JSON.
        /// </summary>
#pragma warning disable IL2026
        public static List<string> DeserializeStringList(string json)
        {
            var context = ExcelToolJsonContext.Default;
            var options = new JsonSerializerOptions 
            { 
                TypeInfoResolver = context,
                AllowTrailingCommas = true,
                ReadCommentHandling = JsonCommentHandling.Skip
            };
            return JsonSerializer.Deserialize<List<string>>(json, options) ?? new List<string>();
        }
#pragma warning restore IL2026
    }
}
