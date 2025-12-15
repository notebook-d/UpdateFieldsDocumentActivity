using Ascon.Pilot.SDK;
using Ascon.Pilot.SDK.Automation;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Windows;
using UpdateFieldsDocumentActivity.Helpers;
using UpdateFieldsDocumentActivity.Message.Views;
using UpdateFieldsDocumentActivity.Services;
using UpdateFieldsDocumentActivity.Settings;
using UpdateFieldsDocumentActivity.Settings.Factories;
using UpdateFieldsDocumentActivity.Settings.Providers;
using System.Runtime.InteropServices.ComTypes;
using DocumentFormat.OpenXml.Drawing;

namespace UpdateFieldsDocumentActivity
{
    [Export(typeof(IAutomationActivity))]
    public class UpdateFieldsDocumentActivity : AutomationActivity
    {
        private readonly IDigitalSigner _digitalSigner;
        private readonly IObjectsRepository _repository;
        private readonly Dictionary<Guid, bool> _processingObjects = new Dictionary<Guid, bool>();
        private readonly object _lockObject = new object();
        
        private readonly int _delayMountFile = 1000;

        private readonly IPersonalSettings _personalSettings;
        private readonly IFileProvider _fileProvider;
        private readonly ISettingProvider<AutoFillSettings> _autoFillGraficFieldsSettings;
        private readonly ISettingProvider<AutoFillSettings> _maskNameTemplateSettings;
        private readonly ISettingProvider<AutoFillSettings> _autoFillBaseSettings;
        private AutoFillSettings _cachedAutoFillGraficFieldsSettings;
        private AutoFillSettings _cachedMaskNameTemplateSettings;
        private AutoFillSettings _cachedAutoFillBaseSettings;
        public const string DOCUMENT_TEMPLATE_FOLDER = "Document_template_folder_793D0CE8-65E6-484E-AAF9-7E095AF9DBD2";
        public const string DOCUMENT_TEMPLATE = "Document_template_89B9E233-A6F9-4B9C-B970-55B3B3A77CED";
        public static readonly Guid DocumentTemplateRootId = new Guid("109DA1F5-7E1A-4DF4-95BD-1FD5AA023DD6");

        [ImportingConstructor]
        public UpdateFieldsDocumentActivity(IDigitalSigner digitalSigner, IObjectsRepository repository, IPersonalSettings personalSettings, IFileProvider fileProvider)
        {
            _digitalSigner = digitalSigner;
            _repository = repository;
            _fileProvider = fileProvider;
            _personalSettings = personalSettings;
            
            _autoFillGraficFieldsSettings = SettingProviderFactory<AutoFillSettings>.GetSettingProvider(personalSettings, SettingsFeatureKeys.AutoFillGraficFieldsSettingsFeatureKey);
            _maskNameTemplateSettings = SettingProviderFactory<AutoFillSettings>.GetSettingProvider(personalSettings, SettingsFeatureKeys.MaskNameTemplateSettingsFeatureKey);
            _autoFillBaseSettings = SettingProviderFactory<AutoFillSettings>.GetSettingProvider(personalSettings, SettingsFeatureKeys.DocsAutoFillFeatureKey); 

            if (Properties.Settings.Default == null 
                || Properties.Settings.Default.IsLogging)
            {
                Logger.EnableLogging();
            }
        }

        public override string Name => nameof(UpdateFieldsDocumentActivity);

        public override Task RunAsync(IObjectModifier modifier, IAutomationBackend backend, IAutomationEventContext context, TriggerType triggerType)
        {
            var objectId = context.Source.Id;
            var objectTypeName = context.Source.Type.Name;
           
            //var taskTemplateRootObject = modifier.EditById(SystemObjectIds.DocumentTemplateRootId).DataObject;
            // Проверяем, не обрабатывается ли уже этот объект
            if (IsProcessing(objectId))
                return Task.CompletedTask;

            //Logger.WriteLog(new LogEntry { Message = $"-----------------------------------------------------------------" });
            //Logger.WriteLog(new LogEntry { Message = $"Обработка объекта {objectTypeName} - {context.Source.DisplayName} ({objectId})" });
            
            if (!context.Source.Attributes.TryGetValue("ImportStatus", out object valueImportStatus))
            {
                //Logger.WriteLog(new LogEntry { Message = $"У объекта {objectId} отсутствует атрибут ImportStatus." });
                //Logger.WriteLog(new LogEntry { Message = $"Обработка объекта {objectId} завершена." });
                return Task.CompletedTask;
            }

            if (!int.TryParse(valueImportStatus.ToString(), out int resultValueImportStatus))
            {
                //Logger.WriteLog(new LogEntry { Message = $"Не удалось преобразовать ImportStatus в число: {valueImportStatus}" });
                //Logger.WriteLog(new LogEntry { Message = $"Обработка объекта {objectId} завершена." });
                return Task.CompletedTask;
            }

            if ((context.Source.Children == null || !context.Source.Children.Any()) 
                && resultValueImportStatus == 1 
                && !AnyObjectHasDocumentFiles(backend, context.Source.Children))
            {
                var latestChild = backend.GetObject(SystemObjectIds.DocumentTemplateRootId)
                                    .Children
                                    .Select(childId => backend.GetObject(childId))
                                    .OrderByDescending(childObj => childObj.Created)
                                    .Where(x=>Convert.ToInt32(x.Attributes[SystemAttributeNames.DOCUMENT_TEMPLATE_TYPE_ID]) == context.Source.Type.Id)
                                    .FirstOrDefault();
                
                if (latestChild != null)
                {     
                    var file = backend.GetObject(latestChild.Children.First());
                    var typeFile = _repository.GetType("File");
                    using (var fileStream = _fileProvider.OpenRead(file.Files.FirstOrDefault()))
                    {
                        var newObj = modifier.Create(context.Source.Id, typeFile);
                        newObj.SetAttribute(SystemAttributeNames.PROJECT_ITEM_NAME, file.DisplayName);
                        newObj.AddFile(file.DisplayName, fileStream, DateTime.Now, DateTime.Now, DateTime.Now);
                    }
                }
            }

            if ((resultValueImportStatus != 1 && resultValueImportStatus != 3) ||
                context.Source.Children == null ||
                !context.Source.Children.Any() ||
                !AnyObjectHasDocumentFiles(backend, context.Source.Children))
            {
                //if (resultValueImportStatus != 1 && resultValueImportStatus != 3)
                //{
                //    Logger.WriteLog(new LogEntry { Message = $"Значение атрибута ImportStatus = {resultValueImportStatus} (не равно 1 или 3)" });
                //}
                //else
                //{
                //    Logger.WriteLog(new LogEntry { Message = $"У объекта {objectId} отсутсвуют дочерние файлы с расширением .DOC, .DOCX, .RTF, .ODT." });
                //}
                
                //Logger.WriteLog(new LogEntry { Message = $"Обработка объекта {objectId} завершена." });
                return Task.CompletedTask;
            }

            Logger.WriteLog(new LogEntry { Message = $"-----------------------------------------------------------------" });
            Logger.WriteLog(new LogEntry { Message = $"Обработка объекта {objectTypeName} - {context.Source.DisplayName} ({objectId})" });
            
            Logger.WriteLog(new LogEntry { Message = $"1 Обновляем значение атрибута ImportStatus" });            
            // Обновляем статус
            UpdateImportStatus(modifier, objectId, resultValueImportStatus);

            // Переименовываем файлы
            List<string> errorIncludePictureList = new List<string>();
            Logger.WriteLog(new LogEntry { Message = $"2 Переименовываем документы согласно маски по шаблону" });
            RenameMaskNameTemplate(modifier, backend, context, objectTypeName, errorIncludePictureList);

            // Запускаем фоновую задачу монтирования файлов
            Logger.WriteLog(new LogEntry { Message = "3 Монтирование файлов на диск" });
            _ = MountFileAsync(context, objectTypeName, errorIncludePictureList);

            return Task.CompletedTask;
        }

        /// <summary>
        /// Кэширование настроек
        /// </summary>
        private AutoFillSettings GetCachedAutoFillGraficFieldsSettings()
        {
            if (_cachedAutoFillGraficFieldsSettings != null)
            {
                return _cachedAutoFillGraficFieldsSettings;
            }
            
            _cachedAutoFillGraficFieldsSettings = _autoFillGraficFieldsSettings.GetSettings();

            return _cachedAutoFillGraficFieldsSettings;
        }

        /// <summary>
        /// Кэширование настроек
        /// </summary>
        private AutoFillSettings GetCachedMaskNameTemplateSettings()
        {
            if (_cachedMaskNameTemplateSettings != null)
            {
                return _cachedMaskNameTemplateSettings;
            }

            _cachedMaskNameTemplateSettings = _maskNameTemplateSettings.GetSettings();

            return _cachedMaskNameTemplateSettings;
        }

        private AutoFillSettings GetCachedAutoFillBaseSettings()
        {
            if (_cachedAutoFillBaseSettings != null)
            {
                return _cachedAutoFillBaseSettings;
            }

            _cachedAutoFillBaseSettings = _autoFillBaseSettings.GetSettings();

            return _cachedAutoFillBaseSettings;
        }

        /// <summary>
        /// Обновляет атрибут статус импорта и запоминает обработанные статусы
        /// </summary>
        private void UpdateImportStatus(IObjectModifier modifier, Guid objectId, int currentStatus)
        {            
            var newImportStatus = currentStatus == 1 ? 2 : 4;
            Logger.WriteLog(new LogEntry { Message = $"Текущее значение атрибута ImportStatus = {currentStatus}" });
            Logger.WriteLog(new LogEntry { Message = $"Присваиваем новое значение атрибуту ImportStatus = {newImportStatus}" });
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
                        var extension = System.IO.Path.GetExtension(file.Name);
                        if (extension != null)
                        {
                            var extensionLower = extension.ToLower();
                            if (Array.Exists(DirectoryService.DOC_EXTENSIONS, ext => ext == extensionLower))
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
        private async Task MountFileAsync(IAutomationEventContext context, string objectTypeName, List<string> errorIncludePictureList)
        {
            var objectId = context.Source.Id;

            // Пытаемся начать обработку
            if (!TryStartProcessing(objectId))
                return;

            // Добавляем задержку 1 секунд перед началом обработки
            await Task.Delay(_delayMountFile);

            try
            {
                Logger.WriteLog(new LogEntry { Message = $"Начало монтирования файлов для объекта {context.Source.DisplayName}" });
                //Проверяем монтированы ли документы на диск
                var pathOld = _repository.GetStoragePath(objectId);
                if (string.IsNullOrEmpty(pathOld))
                {
                    _repository.Mount(objectId);

                    pathOld = await WaitForStoragePathAsync(objectId);
                    if (string.IsNullOrEmpty(pathOld))
                    {
                        Logger.WriteLog(new LogEntry { Message = $"Монтирование файлов на диск", Exception = "Ошибка монтирования файлов на диск." });
                        errorIncludePictureList.Add("Ошибка монтирования файлов на диск.");
                        return;
                    }
                    else
                    {
                        Logger.WriteLog(new LogEntry { Message = $"Файлы смонтированы по пути {pathOld}" });
                    }
                }
                else
                {
                    Logger.WriteLog(new LogEntry { Message = $"Монтирование файлов не требуется. Они смонтированы по пути {pathOld}" });
                }

                //Получаем все файлы word
                var files = DirectoryService.GetWordFiles(pathOld);
                if (files.Count == 0)
                {
                    Logger.WriteLog(new LogEntry { Message = $"По пути {pathOld} не обнаружены файлы с расширением .doc, .docx, .rtf, .odt." });
                    return;
                }

                //Обновляем поля-изображения 
                Logger.WriteLog(new LogEntry { Message = "4 Обновление полей-изображений в документе Word" });
                UpdateAutoFillGraficFields(pathOld, files, objectTypeName, errorIncludePictureList);
                //Обновляем поля
                Logger.WriteLog(new LogEntry { Message = "5 Обновление полей в документе Word" });
                UpdateFieldsDocuments(files, errorIncludePictureList);

                if (errorIncludePictureList.Any())
                {
                    var resultMessage = $"Ошибки:\n{string.Join("\n", errorIncludePictureList)}";

                    _ = Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                    {
                        MessageFactory.CreateDialog(resultMessage, false, "Найдены ошибки").ShowDialog();

                    }), DispatcherPriority.Background);
                }
            }
            finally
            {
                // Всегда освобождаем блокировку, даже при ошибке
                FinishProcessing(objectId);
                errorIncludePictureList.Clear();
                Logger.WriteLog(new LogEntry { Message = $"Обработка объекта {objectId} завершена." });
            }
        }

        /// <summary>
        /// Обновляет путь во всех документах word в полях INCLUDEPICTURE в документе.
        /// </summary>        
        public void UpdateAutoFillGraficFields(string pathFolder, List<string> files, string objectTypeName, List<string> errorIncludePictureList)
        {
            if (string.IsNullOrEmpty(objectTypeName)) return;

            Logger.WriteLog(new LogEntry { Message = "Загрузка настроек 'Автозаполнение графических полей файлов' ..." });
            var autoFillGraficFieldsSettings = GetCachedAutoFillGraficFieldsSettings();
            if (autoFillGraficFieldsSettings?.Types == null) 
            {
                Logger.WriteLog(new LogEntry { Message = "Настройки 'Автозаполнение графических полей файлов' не загружены или не содержат типов" });
                errorIncludePictureList.Add("Настройки 'Автозаполнение графических полей файлов' не загружены или не содержат типов");
                return;
            }

            Logger.WriteLog(new LogEntry { Message = "Настройки 'Автозаполнение графических полей файлов' загружены" });

            var settingsUpdateFieldsCurrentObjectTypes = autoFillGraficFieldsSettings.Types.Where(n => n.Name == objectTypeName && n.Fields != null);
            if (!settingsUpdateFieldsCurrentObjectTypes.Any())
            {
                Logger.WriteLog(new LogEntry { Message = $"Не найдены настройки 'Автозаполнение графических полей файлов' для типа {objectTypeName}" });
                errorIncludePictureList.Add($"Не найдены настройки 'Автозаполнение графических полей файлов' для типа {objectTypeName}");
            }

            foreach (var settingsUpdateFieldsCurrentObjectType in settingsUpdateFieldsCurrentObjectTypes)
            {
                var settingsUpdateFieldsCurrentObjectTypeFields = settingsUpdateFieldsCurrentObjectType.Fields;
                if (!settingsUpdateFieldsCurrentObjectTypeFields.Any())
                {
                    Logger.WriteLog(new LogEntry { Message = $"Не найдены настройки полей для типа {objectTypeName}" });
                    errorIncludePictureList.Add($"Не найдены настройки полей для типа {objectTypeName}");
                }

                foreach (var field in settingsUpdateFieldsCurrentObjectTypeFields)
                {
                    if (string.IsNullOrEmpty(field.PilotFileName)
                        || string.IsNullOrEmpty(field.FileField))
                    {
                        Logger.WriteLog(new LogEntry { Message = $"В настройках для типа {objectTypeName} не заданы поля PilotFileName={field.PilotFileName} или FileField={field.FileField}." });
                        errorIncludePictureList.Add($"В настройках для типа {objectTypeName} не заданы поля PilotFileName={field.PilotFileName} или FileField={field.FileField}.");
                        continue;
                    }
                    else
                    {
                        Logger.WriteLog(new LogEntry { Message = $"Для типа {objectTypeName} PilotFileName={field.PilotFileName} FileField={field.FileField}." });
                    }

                    //Находим картинку по шаблону которую необходимо добавить в поле документа word
                    var latestImageFileSafe = DirectoryService.FindLatestImageFileByPatternSafe(pathFolder, field.PilotFileName);
                    if (string.IsNullOrEmpty(latestImageFileSafe))
                    {
                        Logger.WriteLog(new LogEntry { Message = $"Не удалось найти изображение {field.PilotFileName} по пути {pathFolder}." });
                        errorIncludePictureList.Add($"Не удалось найти изображение {field.PilotFileName} по пути {pathFolder}.");
                        continue;
                    }
                    else
                    {
                        Logger.WriteLog(new LogEntry { Message = $"Найдено изображение {latestImageFileSafe} согласно шаблона {field.PilotFileName} по пути {pathFolder}." });
                    }
                    
                    //Обновляем поля INCLUDEPICTURE
                    foreach (var file in files)
                    {                        
                        try
                        {
                            Logger.WriteLog(new LogEntry { Message = $"Обновление графического поля {field.FileField} в файле {file}" });
                            var resultUpdateIncludePicture = UpdateIncludePictureWithOpenXml.UpdateIncludePictureFieldPath(file, latestImageFileSafe, field.FileField);
                            if (resultUpdateIncludePicture)
                            {
                                Logger.WriteLog(new LogEntry { Message = $"Обновление графического поля {field.FileField} в файле {file} выполено" });
                            }
                            else
                            {
                                var allFieldCodes = UpdateIncludePictureWithOpenXml.CheckIncludePictureFieldPath(file, latestImageFileSafe);
                                if (!allFieldCodes)
                                {
                                    Logger.WriteLog(new LogEntry { Message = $"Обновление графического поля", Exception = $"В документе {file} отсутсвует поле {field.FileField} для вставки графического изображения." });
                                    errorIncludePictureList.Add($"В документе {file} отсутсвует поле {field.FileField} для вставки графического изображения.");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteLog(new LogEntry{ Message = $"Ошибка при обновлении графического поля {field.FileField} в файле {file}: {ex.Message}", Exception = ex.Message });
                        }
                    }
                }
            }            
        }
        
        /// <summary>
        /// Переименовывает файл по маске
        /// </summary>
        /// <param name="pathFolder"></param>
        /// <param name="files"></param>
        /// <param name="objectTypeName"></param>
        public void RenameMaskNameTemplate(IObjectModifier modifier, IAutomationBackend backend, IAutomationEventContext context, string objectTypeName, List<string> errorIncludePictureList)
        {
            if (string.IsNullOrEmpty(objectTypeName)) return;

            Logger.WriteLog(new LogEntry { Message = "Загрузка настроек 'Маски имен документов по шаблону' ..." });
            var maskNameTemplateSettings = GetCachedMaskNameTemplateSettings();
            if (maskNameTemplateSettings?.Types == null)
            {
                Logger.WriteLog(new LogEntry { Message = "Настройки 'Маски имен документов по шаблону' не загружены или не содержат типов" });
                errorIncludePictureList.Add("Настройки 'Маски имен документов по шаблону' не загружены или не содержат типов");
                return;
            }
            Logger.WriteLog(new LogEntry { Message = "Настройки 'Маски имен документов по шаблону' загружены" });

            var settingsUpdateFieldsCurrentObjectTypes = maskNameTemplateSettings.Types.Where(n => n.Name == objectTypeName && n.Fields != null);
            if (!settingsUpdateFieldsCurrentObjectTypes.Any())
            {
                Logger.WriteLog(new LogEntry { Message = $"Не найдены настройки 'Маски имен документов по шаблону' для типа {objectTypeName}" });
                errorIncludePictureList.Add($"Не найдены настройки 'Маски имен документов по шаблону' для типа {objectTypeName}");
            }

            foreach (var settingsUpdateFieldsCurrentObjectType in settingsUpdateFieldsCurrentObjectTypes)
            {
                var settingsUpdateFieldsCurrentObjectTypeFields = settingsUpdateFieldsCurrentObjectType.Fields;
                if (!settingsUpdateFieldsCurrentObjectTypeFields.Any())
                {
                    Logger.WriteLog(new LogEntry { Message = $"Не найдены настройки полей для типа {objectTypeName}" });
                    errorIncludePictureList.Add($"Не найдены настройки полей для типа {objectTypeName}");
                }

                foreach (var field in settingsUpdateFieldsCurrentObjectTypeFields)
                {
                    if (string.IsNullOrEmpty(field.PilotFileName)
                        || string.IsNullOrEmpty(field.NewNameTemplate))
                    {
                        Logger.WriteLog(new LogEntry { Message = $"В настройках для типа {objectTypeName} не заданы поля PilotFileName={field.PilotFileName} или NewNameTemplate={field.NewNameTemplate}." });
                        errorIncludePictureList.Add($"В настройках для типа {objectTypeName} не заданы поля PilotFileName={field.PilotFileName} или NewNameTemplate={field.NewNameTemplate}.");
                        continue;
                    }
                    else
                    {
                        Logger.WriteLog(new LogEntry { Message = $"Для типа {objectTypeName} PilotFileName={field.PilotFileName} NewNameTemplate={field.NewNameTemplate}." });
                    }
                    
                    //Переименовываем документы
                    if (!context.Source.Children.Any())
                    {
                        Logger.WriteLog(new LogEntry { Message = $"У объекта {context.Source.DisplayName} отсутсвуют подчиненные файлы. Переименовывать нечего." });
                    }

                    foreach (var childId in context.Source.Children)
                    {
                        var child = backend.GetObject(childId);
                        var currentFileName = System.IO.Path.GetFileNameWithoutExtension(child.DisplayName);
                        var currentExtFileName = System.IO.Path.GetExtension(child.DisplayName);
                        
                        var tests = DirectoryService.GetTextPartsExcludingCurlyBraces(field.NewNameTemplate);

                        var resTests = DirectoryService.AllTextPartsContainedInTarget(tests, currentFileName);

                        if (currentFileName == field.PilotFileName
                            || resTests)
                        {
                            var nameAttribute = DirectoryService.ExtractValuesFromCurlyBraces(field.NewNameTemplate).FirstOrDefault();
                            //Находим значение по имени атрибута
                            //var atrPilot = context.Source.Attributes.FirstOrDefault(n => n.Key == nameAttribute);
                            if (context.Source.Attributes.TryGetValue(nameAttribute, out object atrPilotValue))
                            {
                                var newNameTemplate = field.NewNameTemplate.Replace($"{{{nameAttribute}}}", atrPilotValue.ToString());
                                var newName = $"{newNameTemplate}{currentExtFileName}".Replace("{", "").Replace("}", "");

                                if (child.DisplayName != newName)
                                {
                                    Logger.WriteLog(new LogEntry { Message = $"Имя файла {child.DisplayName} изменено на {newName}." });
                                    modifier.EditById(child.Id).SetAttribute("$Title", newName);
                                }
                                else
                                {
                                    Logger.WriteLog(new LogEntry { Message = $"Переименование имени файла {child.DisplayName} не требуется." });
                                }
                            }
                            else
                            {
                                Logger.WriteLog(new LogEntry { Message = $"Атрибут {nameAttribute} не найден у объекта {context.Source.DisplayName}" });
                            }
                        }
                        else
                        {
                            Logger.WriteLog(new LogEntry { Message = $"Файл {child.DisplayName} не подходит для переименования по шаблону {field.PilotFileName}." });
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Обновляем поля в документах word
        /// Нужно обновлять даже если нет настроек или ошибки так как возможно графисекие файлы обновлялись
        /// </summary>
        private bool UpdateFieldsDocuments(List<string> files, List<string> errorIncludePictureList)
        {
            Logger.WriteLog(new LogEntry { Message = "Загрузка настроек 'Автозаполнение полей файлов' ..." });
            var autoFillBaseSettings = GetCachedAutoFillBaseSettings();
            if (autoFillBaseSettings?.Types == null)
            {
                Logger.WriteLog(new LogEntry { Message = "Настройки 'Автозаполнение полей файлов' не загружены или не содержат типов" });                
            }
            //Находим все имена полей для word
            var allFileField = Common.GetAllFileFields(autoFillBaseSettings);
            if (!allFileField.Any())
            {
                Logger.WriteLog(new LogEntry { Message = "В настройках 'Автозаполнение полей файлов' не найдены имена полей для обновления в документе" });
                errorIncludePictureList.Add("В настройках 'Автозаполнение полей файлов' не найдены имена полей для обновления в документе");               
            }
            
            foreach (var file in files)
            {
                var lists = UpdateIncludePictureWithOpenXml.GetAllFieldCodes(file);
                foreach (var fileField in allFileField)
                {
                    var isContainsDocPropertyWithName = Common.ContainsDocPropertyWithName(lists, fileField);
                    if (!isContainsDocPropertyWithName)
                    {
                        Logger.WriteLog(new LogEntry { Message = $"В документе {file} отсутствуют поле {fileField} и не может быть заполнена согласно настройке." });
                        errorIncludePictureList.Add($"В документе {file} отсутствуют поле {fileField} и не может быть заполнена согласно настройке.");
                    }
                }
            }
           
            var assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var directoryName = System.IO.Path.GetDirectoryName(assemblyLocation);
            var exeFile = System.IO.Path.Combine(directoryName, @"ShellClientApp\ShellClientApp.exe");

            if (!File.Exists(exeFile))
            {
                Logger.WriteLog(new LogEntry { Message = "Обновление полей", Exception = $"Ошибка: Не удалось найти файл ShellClientApp.exe в расположении {exeFile}" });
                errorIncludePictureList.Add($"Обновление полей. Не удалось найти файл ShellClientApp.exe в расположении {exeFile}");
                return false;
            }

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

                    if (!string.IsNullOrEmpty(output))
                    {
                        Logger.WriteLog(new LogEntry { Message = $"Вызов ShellClientApp для {file}: {output.Replace("\r\n", "")}" });
                    }

                    if (!string.IsNullOrEmpty(error))
                    {
                        Logger.WriteLog(new LogEntry { Message = $"ShellClientApp ошибка для {file}: {error.Replace("\r\n", "")}" });
                        errorIncludePictureList.Add($"ShellClientApp ошибка для {file}: {error.Replace("\r\n", "")}");
                        return false;
                    }
                    else
                    {
                        Logger.WriteLog(new LogEntry { Message = $"В файле {file} поля успешно обновлены" });
                    }
                }
            }

            return true;
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
