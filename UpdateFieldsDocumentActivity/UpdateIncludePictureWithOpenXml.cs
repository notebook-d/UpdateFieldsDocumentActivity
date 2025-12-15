using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using System;
using UpdateFieldsDocumentActivity.Settings;

namespace UpdateFieldsDocumentActivity
{

    public class UpdateIncludePictureWithOpenXml
    {
        /// <summary>
        /// Обновляет путь во всех полях INCLUDEPICTURE в документе.
        /// </summary>
        /// <param name="documentPath">Путь к DOCX-файлу.</param>
        /// <param name="newImagePath">Новый путь к файлу изображения (локальный или сетевой).</param>
        public static bool UpdateIncludePictureFieldPath(string documentPath, string newImagePath, string fieldIdentifierCriterion)
        {
            var result = false;

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

                        result = true;
                    }
                }

                // Сохраняем изменения в документе
                mainPart.Document.Save();
                return result;
            }
        }

        /// <summary>
        /// Признак что поле уже было обновлено
        /// </summary>
        /// <param name="documentPath"></param>
        /// <param name="newImagePath"></param>
        /// <param name="fieldIdentifierCriterion"></param>
        /// <returns></returns>
        public static bool CheckIncludePictureFieldPathOld(string documentPath, string newImagePath)
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
                    //var isTargetField = fieldCode.Parent?.Parent?.InnerText?.Contains(newImagePath) ?? false;
                    var isInclude = IsPathInIncludePicture(currentCode, newImagePath);

                    // Если поле соответствует критериям (или критерий не был задан), обновляем его
                    if (isInclude)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                // Сохраняем изменения в документе
                mainPart.Document.Save();
                return true;
            }
        }
        public static bool CheckIncludePictureFieldPath(string documentPath, string newImagePath)
        {
            // Извлекаем только имя файла из пути
            string fileName = Path.GetFileName(newImagePath);

            if (string.IsNullOrEmpty(fileName))
            {
                return false;
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(documentPath, false))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                // Ищем все элементы FieldCode, содержащие INCLUDEPICTURE
                var fieldCodes = mainPart.Document.Body.Descendants<FieldCode>()
                                             .Where(fc => fc.Text != null && fc.Text.Contains("INCLUDEPICTURE"))
                                             .ToList();

                foreach (var fieldCode in fieldCodes)
                {
                    string currentCode = fieldCode.Text;

                    // Проверяем, содержит ли поле нужное имя файла
                    if (ContainsFileNameInIncludePicture(currentCode, fileName))
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        private static bool ContainsFileNameInIncludePicture(string includePictureText, string fileName)
        {
            if (string.IsNullOrEmpty(includePictureText) || string.IsNullOrEmpty(fileName))
                return false;

            // Извлекаем путь из INCLUDEPICTURE
            string pathFromField = ExtractPathFromIncludePicture(includePictureText);

            if (string.IsNullOrEmpty(pathFromField))
                return false;

            // Извлекаем имя файла из пути в поле
            string fieldFileName = Path.GetFileName(pathFromField);

            if (string.IsNullOrEmpty(fieldFileName))
                return false;

            // Сравниваем имена файлов без учета регистра
            return string.Equals(fieldFileName, fileName, StringComparison.OrdinalIgnoreCase);
        }

        private static string ExtractPathFromIncludePicture(string includePictureText)
        {
            // Ищем путь в кавычках после INCLUDEPICTURE
            int startQuote = includePictureText.IndexOf('"');
            int endQuote = includePictureText.LastIndexOf('"');

            if (startQuote >= 0 && endQuote > startQuote)
            {
                return includePictureText.Substring(startQuote + 1, endQuote - startQuote - 1);
            }

            return string.Empty;
        }


        private static string ExtractPathFromIncludePictureOld(string includePictureText)
        {
            // Ищем путь в кавычках
            int startQuote = includePictureText.IndexOf('"');
            int endQuote = includePictureText.LastIndexOf('"');

            if (startQuote >= 0 && endQuote > startQuote)
            {
                return includePictureText.Substring(startQuote + 1, endQuote - startQuote - 1);
            }

            // Альтернативный способ через Regex
            var match = Regex.Match(includePictureText, @"INCLUDEPICTURE\s+""(.*?)""");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            return string.Empty;
        }

        public static bool IsPathInIncludePicture(string includePictureText, string filePath)
        {
            if (string.IsNullOrEmpty(includePictureText) || string.IsNullOrEmpty(filePath))
                return false;

            // Извлекаем путь из INCLUDEPICTURE
            string extractedPath = ExtractPathFromIncludePicture(includePictureText);
            if (string.IsNullOrEmpty(extractedPath))
                return false;

            // Нормализуем оба пути
            string normalizedFilePath = NormalizePath(filePath);
            string normalizedExtractedPath = NormalizePath(extractedPath);

            // Сравниваем
            return string.Equals(normalizedFilePath, normalizedExtractedPath,
                StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            // 1. Нормализуем Unicode символы (решает проблему с тильдой)
            path = path.Normalize(NormalizationForm.FormC);

            // 2. Заменяем обратные слеши на прямые
            path = path.Replace('\\', '/');

            // 3. Убираем кавычки
            path = path.Trim('"');

            // 4. Убираем лишние пробелы
            path = path.Trim();

            return path;
        }


        public static List<string> GetAllFieldCodes(string documentPath)
        {
            var allFieldCodes = new List<string>();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(documentPath, false))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                var body = mainPart.Document.Body;

                // 1. Собираем все FieldCode
                var fieldCodeElements = body.Descendants<FieldCode>().ToList();
                foreach (var fieldCode in fieldCodeElements)
                {
                    string code = fieldCode.Text?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(code))
                    {
                        allFieldCodes.Add(code);
                    }
                }

                // 2. Собираем SimpleField
                var simpleFields = body.Descendants<SimpleField>().ToList();
                foreach (var simpleField in simpleFields)
                {
                    string instruction = simpleField.Instruction?.Value ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(instruction))
                    {
                        allFieldCodes.Add(instruction);
                    }
                }

                // 3. Ищем дополнительные поля
                var fieldChars = body.Descendants<FieldChar>()
                    .Where(fc => fc.FieldCharType?.Value == FieldCharValues.Begin)
                    .ToList();

                foreach (var fieldChar in fieldChars)
                {
                    var fieldCode = FindFieldCodeForFieldChar(fieldChar);
                    if (fieldCode != null && !string.IsNullOrWhiteSpace(fieldCode.Text))
                    {
                        // Проверяем, не добавили ли мы уже это поле
                        string code = fieldCode.Text.Trim();
                        if (!allFieldCodes.Contains(code))
                        {
                            allFieldCodes.Add(code);
                        }
                    }
                }

                // 4. Проверяем текст в Run на наличие полей в фигурных скобках
                var runs = body.Descendants<Run>().ToList();
                foreach (var run in runs)
                {
                    var textElements = run.Descendants<Text>();
                    foreach (var text in textElements)
                    {
                        string textValue = text.Text ?? string.Empty;
                        if (textValue.Contains("{") && textValue.Contains("}"))
                        {
                            // Пытаемся извлечь поле из фигурных скобок
                            int start = textValue.IndexOf('{');
                            int end = textValue.IndexOf('}', start);
                            if (start >= 0 && end > start)
                            {
                                string fieldText = textValue.Substring(start + 1, end - start - 1).Trim();
                                if (!string.IsNullOrWhiteSpace(fieldText))
                                {
                                    allFieldCodes.Add(fieldText);
                                }
                            }
                        }
                    }
                }
            }

            return allFieldCodes.Distinct().ToList();
        }

        private static FieldCode FindFieldCodeForFieldChar(FieldChar fieldChar)
        {
            // Ищем FieldCode в том же Run или в предыдущих Run
            var currentRun = fieldChar.Parent as Run;

            // Проверяем текущий Run
            var fieldCodeInCurrentRun = currentRun?.Elements<FieldCode>().FirstOrDefault();
            if (fieldCodeInCurrentRun != null)
                return fieldCodeInCurrentRun;

            // Ищем в предыдущих Run
            var previousRun = currentRun?.PreviousSibling() as Run;
            while (previousRun != null)
            {
                var fieldCode = previousRun.Elements<FieldCode>().FirstOrDefault();
                if (fieldCode != null)
                    return fieldCode;

                previousRun = previousRun.PreviousSibling() as Run;
            }

            return null;
        }

        
    }
}
