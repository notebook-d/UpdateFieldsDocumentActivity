using System;
using System.IO;
using System.Linq;


namespace UpdateFieldsDocumentActivity.Services
{
    public class DirectoryService
    {
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
                throw new DirectoryNotFoundException($"Директория не найдена: {directoryPath}");
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
                Console.WriteLine($"Ошибка: {ex.Message}");
                return null;
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"Нет доступа к директории: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка при поиске файла: {ex.Message}");
                return null;
            }
        }
    }
}
