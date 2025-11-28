using Ascon.Pilot.SDK;
using Ascon.Pilot.SDK.Automation;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Threading;
using System.Threading.Tasks;
using UpdateFieldsDocumentActivity.Services;
using UpdateFieldsDocumentActivity.Settings;
using UpdateFieldsDocumentActivity.Settings.Factories;
using UpdateFieldsDocumentActivity.Settings.Providers;

namespace UpdateFieldsDocumentActivity
{
    [Export(typeof(IAutomationActivity))]
    public class UpdateFieldsDocumentActivity : AutomationActivity
    {
        private readonly IDigitalSigner _digitalSigner;
        private readonly IObjectsRepository _repository;
        private readonly Dictionary<Guid, bool> _processingObjects = new Dictionary<Guid, bool>();
        private readonly object _lockObject = new object();

        private static readonly string[] _docExtensions = { ".doc", ".docx", ".rtf", ".odt" };
        private readonly int _delayMountFile = 1000;

        private readonly AutoFillSettings _autoFillSettings;

        private readonly IPersonalSettings _personalSettings;        
        private readonly ISettingProvider<AutoFillSettings> _settingsUpdateFieldsDocumentIncludePicture;

        [ImportingConstructor]
        public UpdateFieldsDocumentActivity(IDigitalSigner digitalSigner, IObjectsRepository repository, IPersonalSettings personalSettings)
        {
            _digitalSigner = digitalSigner;
            _repository = repository;

            _personalSettings = personalSettings;
            _settingsUpdateFieldsDocumentIncludePicture = SettingProviderFactory<AutoFillSettings>.GetSettingProvider(personalSettings, SettingsFeatureKeys.SettingKey);
        }

        public override string Name => nameof(UpdateFieldsDocumentActivity);

        public override Task RunAsync(IObjectModifier modifier, IAutomationBackend backend, IAutomationEventContext context, TriggerType triggerType)
        {
            var objectId = context.Source.Id;
            var objectTypeName = context.Source.Type.Name;

            // Проверяем, не обрабатывается ли уже этот объект
            if (IsProcessing(objectId))
                return Task.CompletedTask;

            if (!context.Source.Attributes.TryGetValue("ImportStatus", out object valueImportStatus))
                return Task.CompletedTask;

            if (!int.TryParse(valueImportStatus.ToString(), out int resultValueImportStatus))
                return Task.CompletedTask;

            if ((resultValueImportStatus != 1 && resultValueImportStatus != 3) ||
                context.Source.Children == null ||
                !context.Source.Children.Any() ||
                !AnyObjectHasDocumentFiles(backend, context.Source.Children))
            {
                return Task.CompletedTask;
            }

            UpdateImportStatus(modifier, context, objectId, resultValueImportStatus);

            _ = MountFileAsync(context, objectTypeName);

            return Task.CompletedTask;
        }

        /// <summary>
        /// Обновляет атрибут статус импорта и запоминает обработанные статусы
        /// </summary>
        private void UpdateImportStatus(IObjectModifier modifier, IAutomationEventContext context, Guid objectId, int currentStatus)
        {
            var newImportStatus = currentStatus == 1 ? 2 : 4;
            modifier.EditById(objectId).SetAttribute("ImportStatus", newImportStatus);
        }

        /// <summary>
        /// Признак наличия документов типа .DOC, .DOCX, .RTF и .ODT
        /// </summary>
        private bool AnyObjectHasDocumentFiles(IAutomationBackend backend, ReadOnlyCollection<Guid> childIds)
        {
            foreach (var idChildrenObject in childIds)
            {
                var childrenObject = backend.GetObject(idChildrenObject);

                if (childrenObject.Files != null)
                {
                    foreach (var file in childrenObject.Files)
                    {
                        var extension = Path.GetExtension(file.Name);
                        if (extension != null)
                        {
                            var extensionLower = extension.ToLower();
                            if (Array.Exists(_docExtensions, ext => ext == extensionLower))
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Монтирует файлы на диск
        /// </summary>
        private async Task MountFileAsync(IAutomationEventContext context, string objectTypeName)
        {
            var objectId = context.Source.Id;

            // Пытаемся начать обработку
            if (!TryStartProcessing(objectId))
                return;

            // Добавляем задержку 1 секунд перед началом обработки
            await Task.Delay(_delayMountFile);

            try
            {
                //Проверяем монтированы ли документы на диск
                var pathOld = _repository.GetStoragePath(objectId);
                if (string.IsNullOrEmpty(pathOld))
                {
                    _repository.Mount(objectId);

                    pathOld = await WaitForStoragePathAsync(objectId);
                    if (string.IsNullOrEmpty(pathOld))
                        return;
                }

                var files = GetWordFiles(pathOld);
                if (files.Count == 0) 
                    return;

                //Обновляем поля-изображения 
                UpdateIncludePictureField(pathOld, files, objectTypeName);
                //Обновляем поля
                UpdateFieldsDocuments(pathOld, files);
            }
            finally
            {
                // Всегда освобождаем блокировку, даже при ошибке
                FinishProcessing(objectId);
            }
        }

        /// <summary>
        /// Обновляет путь во всех докумннтах word согласно  полях INCLUDEPICTURE в документе.
        /// </summary>
        /// <param name="documentPath"></param>
        /// <param name="newImagePath"></param>
        /// <param name="settingsUpdateFieldsDocumentIncludePicture"></param>
        public void UpdateIncludePictureField(string pathFolder, List<string> files, string objectTypeName)
        {
            if (string.IsNullOrEmpty(objectTypeName)) return;

            var settingsUpdateFieldsDocumentIncludePicture = _settingsUpdateFieldsDocumentIncludePicture.GetSettings();
            if (settingsUpdateFieldsDocumentIncludePicture?.Types == null) return;

            var settingsUpdateFieldsCurrentObjectTypes = settingsUpdateFieldsDocumentIncludePicture.Types.Where(n => n.Name == objectTypeName);
            foreach (var settingsUpdateFieldsCurrentObjectType in settingsUpdateFieldsCurrentObjectTypes)
            {
                if (settingsUpdateFieldsCurrentObjectType.Fields == null) return;
                foreach (var field in settingsUpdateFieldsCurrentObjectType.Fields)
                {
                    if (string.IsNullOrEmpty(field.PilotFileName)
                        || string.IsNullOrEmpty(field.FileField)) continue;

                    //Находим картинку по шаблону которую необходимо добавить в поле документа word
                    var latestImageFileSafe = DirectoryService.FindLatestImageFileByPatternSafe(pathFolder, field.PilotFileName);
                    if (string.IsNullOrEmpty(latestImageFileSafe)) continue;
                    
                    //Обновляем поля INCLUDEPICTURE
                    foreach (var file in files)
                    {
                        UpdateIncludePictureWithOpenXml.UpdateIncludePictureFieldPath(file, latestImageFileSafe, field.FileField);
                    }
                }
            }
            return;
        }

        /// <summary>
        /// Обновляем поля в документах word
        /// </summary>
        private bool UpdateFieldsDocuments(string pathFolder, List<string> files)
        {
            var assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var directoryName = Path.GetDirectoryName(assemblyLocation);
            var exeFile = Path.Combine(directoryName, @"ShellClientApp\ShellClientApp.exe");

            if (!File.Exists(exeFile))
                throw new Exception(string.Format("Ошибка: Не удалось найти файл ShellClientApp.exe в расположении {0}", exeFile));

            //var files = GetWordFiles(pathFolder);

            foreach (var file in files)
            {
                var processStartInfo = new ProcessStartInfo
                {
                    FileName = exeFile,
                    Arguments = string.Format("\"{0}\"", file),
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = Process.Start(processStartInfo))
                {
                    process.WaitForExit();

                    var output = process.StandardOutput.ReadToEnd();
                    var error = process.StandardError.ReadToEnd();

                    if (!string.IsNullOrEmpty(error))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Возвращает все файлы word с диска
        /// </summary>
        private static List<string> GetWordFiles(string folderPath)
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
                    if (Array.Exists(_docExtensions, ext => ext == extensionLower))
                    {
                        result.Add(file);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Ожидание монтирования файлов
        /// </summary>
        private async Task<string> WaitForStoragePathAsync(Guid id, int timeout = 5000)
        {
            var cts = new CancellationTokenSource(timeout);

            while (!cts.Token.IsCancellationRequested)
            {
                var path = _repository.GetStoragePath(id);
                if (!string.IsNullOrEmpty(path))
                    return path;

                await Task.Delay(100, cts.Token);
            }

            return null;
        }

        #region ProcessingMountFile

        private bool IsProcessing(Guid objectId)
        {
            lock (_lockObject)
            {
                return _processingObjects.ContainsKey(objectId);
            }
        }

        private bool TryStartProcessing(Guid objectId)
        {
            lock (_lockObject)
            {
                if (_processingObjects.ContainsKey(objectId))
                    return false;

                _processingObjects[objectId] = true;
                return true;
            }
        }

        private void FinishProcessing(Guid objectId)
        {
            lock (_lockObject)
            {
                _processingObjects.Remove(objectId);
            }
        }

        #endregion

    }
}
