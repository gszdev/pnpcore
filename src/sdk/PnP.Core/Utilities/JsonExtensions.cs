using System;
using System.Text.Json;

namespace PnP.Core
{
    /// <summary>
    /// Conversion to object extensions
    /// </summary>
    public static class JsonExtensions
    {
        /// <summary>
        /// Deserializes a JsonElement to an Object
        /// </summary>
        public static T ToObject<T>(this JsonElement element, JsonSerializerOptions options = null)
        {
            var json = element.GetRawText();
            return JsonSerializer.Deserialize<T>(json, options);
        }

        /// <summary>
        /// Deserializes a JsonDocument to an Object
        /// </summary>
        public static T ToObject<T>(this JsonDocument document, JsonSerializerOptions options = null)
        {
            if (document == null)
            {
                throw new ArgumentException("Document is null", nameof(document));
            }
            var json = document.RootElement.GetRawText();
            return JsonSerializer.Deserialize<T>(json, options);
        }
    }

}