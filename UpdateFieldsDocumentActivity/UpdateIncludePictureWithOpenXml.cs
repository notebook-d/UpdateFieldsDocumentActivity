using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.IO;

namespace UpdateFieldsDocumentActivity
{

    public class UpdateIncludePictureWithOpenXml
    {
        /// <summary>
        /// Обновляет путь во всех полях INCLUDEPICTURE в документе.
        /// </summary>
        /// <param name="documentPath">Путь к DOCX-файлу.</param>
        /// <param name="newImagePath">Новый путь к файлу изображения (локальный или сетевой).</param>
        public static void UpdateIncludePictureFieldPath(string documentPath, string newImagePath, string fieldIdentifierCriterion)
        {
            // Убедитесь, что новый путь использует правильные разделители для XML (обратный слеш обычно работает)
            string normalizedNewImagePath = newImagePath.Replace('\\', '/');

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(documentPath, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                // Ищем все элементы FieldCode, содержащие INCLUDEPICTURE
                var fieldCodes = mainPart.Document.Body.Descendants<FieldCode>()
                                             .Where(fc => fc.Text.Contains("INCLUDEPICTURE"))
                                             .ToList();

                foreach (var fieldCode in fieldCodes)
                {
                    string currentCode = fieldCode.Text;
                    var isTargetField = fieldCode.Parent?.Parent?.InnerText?.Contains(fieldIdentifierCriterion) ?? false;

                    // Если поле соответствует критериям (или критерий не был задан), обновляем его
                    if (isTargetField)
                    {
                        // Используем Regex для извлечения *всего*, что находится после закрывающей кавычки пути,
                        // чтобы сохранить существующие ключи (\d, \* MERGEFORMAT) и наш комментарий /*...*/
                        var switchesMatch = Regex.Match(currentCode, @"INCLUDEPICTURE\s+"".*?""\s*(.*)");
                        string existingSwitchesAndComments = switchesMatch.Success ? switchesMatch.Groups[1].Value.Trim() : string.Empty;

                        // Формируем полностью новый код поля
                        string updatedCode = $"INCLUDEPICTURE \"{normalizedNewImagePath}\" {existingSwitchesAndComments}".Trim();

                        // Заменяем старый код новым
                        fieldCode.Text = updatedCode;
                    }
                }

                // Сохраняем изменения в документе
                mainPart.Document.Save();
            }
        }
    }
}
