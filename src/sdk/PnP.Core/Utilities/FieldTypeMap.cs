using AngleSharp.Dom;
using AngleSharp.Io;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PnP.Core.Utilities
{
    internal static class FieldTypeMap
    {
        private static readonly Lazy<IDictionary<string, string>> mappings = new Lazy<IDictionary<string, string>>(BuildMappings);

        private static IDictionary<string, string> BuildMappings()
        {
            var mappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Note"] = "SP.FieldMultiLineText",
                ["Choice"] = "SP.FieldChoice",
                ["MultiChoice"] = "SP.FieldMultiChoice",
                ["Number"] = "SP.FieldNumber",
                ["Currency"] = "SP.FieldCurrency",
                ["DateTime"] = "SP.FieldDateTime",
                ["Integer"] = "SP.FieldNumber",

                ["Lookup"] = "SP.FieldLookup",
                ["LookupMulti"] = "SP.FieldLookup",

                ["User"] = "SP.FieldUser",
                ["UserMulti"] = "SP.FieldUser",

                ["URL"] = "SP.FieldUrl",
                ["Calculated"] = "SP.FieldCalculated",

                ["TaxonomyFieldType"] = "SP.Taxonomy.TaxonomyField",
                ["TaxonomyFieldTypeMulti"] = "SP.Taxonomy.TaxonomyField",

                ["ContentTypeId"] = "SP.FieldContentTypeId",
                ["Guid"] = "SP.FieldGuid",
                ["Computed"] = "SP.FieldComputed",
            };

            /*
            var cache = mappings.ToList(); // need ToList() to avoid modifying while still enumerating

            foreach (var mapping in cache)
            {
                if (!mappings.ContainsKey(mapping.Value))
                {
                    mappings.Add(mapping.Value, mapping.Key);
                }
            }
            */

            return mappings;
        }

        /// <summary>
        /// Tries to get the type of the entity from the provided string.
        /// </summary>
        /// <param name="typeAsString">The field type as string.</param>
        /// <param name="entityType">The variable to store the entity type.</param>
        /// <returns>The entity type.</returns>
        /// <exception cref="ArgumentNullException" />
        internal static bool TryGetFieldEntityType(string typeAsString, out string entityType)
        {
            if (typeAsString == null)
            {
                throw new ArgumentNullException(nameof(typeAsString));
            }

            return mappings.Value.TryGetValue(typeAsString, out entityType);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="typeAsString"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static string GetByTypeAsString(string typeAsString, string defaultValue = "SP.Field")
        {
            if (typeAsString == null)
            {
                throw new ArgumentNullException(nameof(typeAsString));
            }

            if (TryGetFieldEntityType(typeAsString, out var entityType))
            {
                return entityType;
            }

            return defaultValue;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="contextAwareObject"></param>
        /// <param name="entity"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static string TryToGetConcretFieldMetadataType(IDataModelWithContext contextAwareObject, EntityInfo entity, string defaultValue = "SP.Field")
        {
            if (contextAwareObject == null)
            {
                throw new ArgumentNullException(nameof(contextAwareObject));
            }

            if (entity == null)
            {
                throw new ArgumentNullException(nameof(entity));
            }

            if (contextAwareObject is Field field
                && !string.IsNullOrEmpty(field.TypeAsString)
                && TryGetFieldEntityType(field.TypeAsString, out var entityType))
            {
                return entityType;
            }

            return entity.SharePointType;
        }
    }
}
