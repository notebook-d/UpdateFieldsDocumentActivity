using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UpdateFieldsDocumentActivity.Settings;

namespace UpdateFieldsDocumentActivity.Services
{
    internal class Common
    {
        public static bool ContainsDocPropertyWithName(List<string> strings, string fieldName)
        {
            if (strings == null || strings.Count == 0)
                return false;

            foreach (string str in strings)
            {
                if (!string.IsNullOrEmpty(str) &&
                    str.IndexOf("DOCPROPERTY", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    str.IndexOf(fieldName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static List<string> GetAllFileFields(AutoFillSettings settings)
        {
            if (settings?.Types == null)
                return new List<string>();

            var fileFields = new List<string>();

            foreach (var docType in settings.Types)
            {
                // Получаем поля из старого формата
                if (docType.Fields != null)
                {
                    foreach (var field in docType.Fields)
                    {
                        if (!string.IsNullOrWhiteSpace(field.FileField))
                        {
                            fileFields.Add(field.FileField.Trim());
                        }
                    }
                }

                // Получаем поля из нового формата
                if (docType.FieldsContainer?.FieldMappings != null)
                {
                    foreach (var field in docType.FieldsContainer.FieldMappings)
                    {
                        if (!string.IsNullOrWhiteSpace(field.FileField))
                        {
                            fileFields.Add(field.FileField.Trim());
                        }
                    }
                }
            }

            return fileFields.Distinct().OrderBy(f => f).ToList();
        }
    }
}
