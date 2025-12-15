using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using UpdateFieldsDocumentActivity.Helpers;


namespace UpdateFieldsDocumentActivity.Services
{
    public class DirectoryService
    {
        public static readonly string[] DOC_EXTENSIONS = { ".doc", ".docx", ".rtf", ".odt" };

        /// <summary>
        /// Находит самый свежий графический файл по одному шаблону
        /// </summary>
        /// <param name="directoryPath">Путь к директории для поиска</param>
        /// <param name="searchPattern">Шаблон поиска имени файла (например, "*схема*", "*диаграмма*", "*план*")</param>
        /// <returns>Полный путь к самому свежему графическому файлу или null</returns>
        public static string FindLatestImageFileByPattern(string directoryPath, string searchPattern)
        {
            // Проверяем существование директории
            if (!Directory.Exists(directoryPath))
            {
                Logger.WriteLog(new LogEntry { Message = $"Поиск изображения {searchPattern} по пути {directoryPath}.", Exception = $"Директория не найдена: {directoryPath}" });
                return null;
            }

            // Расширения графических файлов
            string[] imageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.gif", "*.tiff", "*.webp" };

            // Ищем все графические файлы по шаблону и выбираем самый свежий
            var latestImage = imageExtensions
                .SelectMany(ext => Directory.GetFiles(directoryPath, $"{searchPattern}{ext.Replace("*", "")}"))
                .OrderByDescending(file => new FileInfo(file).LastWriteTime)
                .FirstOrDefault();

            return latestImage;
        }

        /// <summary>
        /// Находит самый свежий графический файл по одному шаблону с безопасной обработкой исключений
        /// </summary>
        /// <param name="directoryPath">Путь к директории для поиска</param>
        /// <param name="searchPattern">Шаблон поиска имени файла</param>
        /// <returns>Полный путь к самому свежему графическому файлу или null</returns>
        public static string FindLatestImageFileByPatternSafe(string directoryPath, string searchPattern)
        {
            try
            {
                return FindLatestImageFileByPattern(directoryPath, searchPattern);
            }
            catch (DirectoryNotFoundException ex)
            {
                Logger.WriteLog(new LogEntry { Message = $"Поиск изображения {searchPattern} по пути {directoryPath}.", Exception=$"{ex.Message}" });
                return null;
            }
            catch (UnauthorizedAccessException ex)
            {
                Logger.WriteLog(new LogEntry { Message = $"Поиск изображения {searchPattern} по пути {directoryPath}.", Exception = $"Нет доступа к директории: {ex.Message}" });
                return null;
            }
            catch (Exception ex)
            {
                Logger.WriteLog(new LogEntry { Message = $"Поиск изображения {searchPattern} по пути {directoryPath}.", Exception = $"Произошла ошибка при поиске файла: {ex.Message}" });
                return null;
            }
        }

        /// <summary>
        /// Находит все файлы по шаблону nameTemplate
        /// </summary>
        /// <param name="directoryPath"></param>
        /// <param name="nameTemplate"></param>
        /// <returns></returns>
        /// <exception cref="DirectoryNotFoundException"></exception>
        public static string[] FindFilesByNameTemplate(string directoryPath, string nameTemplate)
        {
            if (!Directory.Exists(directoryPath))
                throw new DirectoryNotFoundException($"Директория не найдена: {directoryPath}");

            // Используем SearchOption.AllDirectories для поиска во всех подпапках
            // или SearchOption.TopDirectoryOnly только в текущей папке
            return Directory.GetFiles(directoryPath, nameTemplate, SearchOption.AllDirectories);
        }

        /// <summary>
        /// Находит все файлы с именем nameTemplate
        /// </summary>
        /// <param name="directoryPath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="DirectoryNotFoundException"></exception>
        public static string[] FindFilesExactName(string directoryPath, string fileName)
        {
            if (!Directory.Exists(directoryPath))
                throw new DirectoryNotFoundException($"Директория не найдена: {directoryPath}");

            // Ищем все файлы и фильтруем по точному имени (без расширения)
            return Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories)
                           .Where(file => Path.GetFileNameWithoutExtension(file).Equals(fileName, StringComparison.OrdinalIgnoreCase))
                           .ToArray();
        }

        /// <summary>
        /// Возвращает все файлы word с диска
        /// </summary>
        public static List<string> GetWordFiles(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath))
                return new List<string>();

            if (!Directory.Exists(folderPath))
                return new List<string>();

            var allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);
            var result = new List<string>();

            foreach (var file in allFiles)
            {
                var extension = Path.GetExtension(file);
                if (extension != null)
                {
                    var extensionLower = extension.ToLower();
                    if (Array.Exists(DOC_EXTENSIONS, ext => ext == extensionLower))
                    {
                        result.Add(file);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Находит все значения в фигурных скобках в тексте
        /// </summary>
        /// <param name="text">Исходный текст</param>
        /// <returns>Список найденных значений</returns>
        public static List<string> ExtractValuesFromCurlyBraces(string text)
        {
            if (string.IsNullOrEmpty(text))
                return new List<string>();

            var matches = Regex.Matches(text, @"\{([^{}]+)\}");
            return matches.Cast<Match>()
                         .Select(m => m.Groups[1].Value)
                         .ToList();
        }

        /// <summary>
        /// Возвращает все части текста кроме содержимого в фигурных скобках
        /// </summary>
        public static List<string> GetTextPartsExcludingCurlyBraces(string text)
        {
            if (string.IsNullOrEmpty(text))
                return new List<string>();

            // Разделяем текст по шаблону фигурных скобок
            var parts = Regex.Split(text, @"\{[^{}]+\}")
                            .Where(part => !string.IsNullOrWhiteSpace(part))
                            .ToList();

            return parts;
        }

        /// <summary>
        /// Проверяет, все ли части текста содержатся в целевой строке
        /// </summary>
        public static bool AllTextPartsContainedInTarget(List<string> textParts, string targetString)
        {
            if (textParts == null || textParts.Count == 0)
                return true; // Нет частей для проверки - считаем валидным

            if (string.IsNullOrEmpty(targetString))
                return false; // Целевая строка пустая, но есть части для проверки

            return textParts.All(part => targetString.Contains(part));
        }

    }
}
