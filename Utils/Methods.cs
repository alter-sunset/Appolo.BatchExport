using Appolo.BatchExport.Utils;
using Appolo.Utilities;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Appolo.BatchExport
{
    public static class Methods
    {
        public static void ExportModelToNWC(Document document, NWCExportUi ui, ref bool isFuckedUp, Logger logger)
        {
            Element view = default;

            //Doesn't help due to immediate error between nwc(nwd) links and once opened navisExportModule
            //DeleteCoordinationModels(document);            

            using (FilteredElementCollector stuff = new FilteredElementCollector(document))
            {
                view = stuff.OfClass(typeof(View3D)).FirstOrDefault(e => e.Name == ui.TextBoxExportScopeViewName.Text && !((View3D)e).IsTemplate);
            }

            if ((bool)ui.RadioButtonExportScopeView.IsChecked
                && !(bool)ui.CheckBoxExportLinks.IsChecked
                && IsViewEmpty(document, view))
            {
                logger.Error("Нет геометрии на виде.");
                isFuckedUp = true;
            }
            else
            {
                NavisworksExportOptions navisworksExportOptions = ExportModelToNWCOptions(document, ui);
                string folder = "";
                string prefix = "";
                string postfix = "";

                ui.Dispatcher.Invoke(() => folder = ui.TextBoxFolder.Text);
                ui.Dispatcher.Invoke(() => prefix = ui.TextBoxPrefix.Text);
                ui.Dispatcher.Invoke(() => postfix = ui.TextBoxPostfix.Text);

                string fileExportName = prefix + document.Title.Replace("_отсоединено", "") + postfix;
                string fileName = folder + "\\" + fileExportName + ".nwc";

                string oldHash = null;

                if (File.Exists(fileName))
                {
                    oldHash = MD5Hash(fileName);
                    logger.Hash(oldHash);
                }

                try
                {
                    document.Export(folder, fileExportName, navisworksExportOptions);
                }
                catch (Exception ex)
                {
                    logger.Error("Смотри исключение.", ex);
                    isFuckedUp = true;
                }

                navisworksExportOptions.Dispose();

                if (!File.Exists(fileName))
                {
                    logger.Error("Файл не был создан. Скорее всего нет геометрии на виде.");
                    isFuckedUp = true;
                }
                else
                {
                    string newHash = MD5Hash(fileName);
                    logger.Hash(newHash);

                    if (newHash == oldHash)
                    {
                        logger.Error("Файл не был обновлён. Хэш сумма не изменилась.");
                        isFuckedUp = true;
                    }
                }

                if (!(view is null))
                {
                    view.Dispose();
                }
            }
        }
        public static void ExportModelToIFC(Document document, IFCExportUi ui, ref bool isFuckedUp, Logger logger)
        {
            Element view = default;

            //Doesn't help due to immediate error between nwc(nwd) links and once opened navisExportModule
            //DeleteCoordinationModels(document);            

            using (FilteredElementCollector stuff = new FilteredElementCollector(document))
            {
                view = stuff.OfClass(typeof(View3D)).FirstOrDefault(e => e.Name == ui.TextBoxExportScopeViewName.Text && !((View3D)e).IsTemplate);
            }

            if ((bool)ui.RadioButtonExportScopeView.IsChecked
                && IsViewEmpty(document, view))
            {
                logger.Error("Нет геометрии на виде.");
                isFuckedUp = true;
            }
            else
            {
                IFCExportOptions iFCExportOptions = new IFCExportOptions();
                string folder = "";
                string prefix = "";
                string postfix = "";

                ui.Dispatcher.Invoke(() => folder = ui.TextBoxFolder.Text);
                ui.Dispatcher.Invoke(() => prefix = ui.TextBoxPrefix.Text);
                ui.Dispatcher.Invoke(() => postfix = ui.TextBoxPostfix.Text);

                string fileExportName = prefix + document.Title.Replace("_отсоединено", "") + postfix;
                string fileName = folder + "\\" + fileExportName + ".ifc";

                string oldHash = null;

                if (File.Exists(fileName))
                {
                    oldHash = MD5Hash(fileName);
                    logger.Hash(oldHash);
                }

                using (Transaction transaction = new Transaction(document))
                {
                    transaction.Start("Экспорт IFC");

                    try
                    {
                        document.Export(folder, fileExportName, iFCExportOptions);
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Смотри исключение.", ex);
                        isFuckedUp = true;
                    }
                    transaction.Commit();
                }


                iFCExportOptions.Dispose();

                if (!File.Exists(fileName))
                {
                    logger.Error("Файл не был создан. Скорее всего нет геометрии на виде.");
                    isFuckedUp = true;
                }
                else
                {
                    string newHash = MD5Hash(fileName);
                    logger.Hash(newHash);

                    if (newHash == oldHash)
                    {
                        logger.Error("Файл не был обновлён. Хэш сумма не изменилась.");
                        isFuckedUp = true;
                    }
                }

                if (!(view is null))
                {
                    view.Dispose();
                }
            }
        }

        public static NavisworksExportOptions ExportModelToNWCOptions(Document document, NWCExportUi batchExportNWC)
        {
            string coordinates = ((ComboBoxItem)batchExportNWC.ComboBoxCoordinates.SelectedItem).Content.ToString();
            bool exportScope = (bool)batchExportNWC.RadioBattonExportScopeModel.IsChecked;
            string parameters = ((ComboBoxItem)batchExportNWC.ComboBoxParameters.SelectedItem).Content.ToString();

            if (double.TryParse(batchExportNWC.TextBoxFacetingFactor.Text, out double facetingFactor))
            {
                facetingFactor = double.Parse(batchExportNWC.TextBoxFacetingFactor.Text);
            }
            else
            {
                facetingFactor = 1.0;
            }

            NavisworksExportOptions options = new NavisworksExportOptions()
            {
                ConvertElementProperties = (bool)batchExportNWC.CheckBoxConvertElementProperties.IsChecked,
                DivideFileIntoLevels = (bool)batchExportNWC.CheckBoxDivideFileIntoLevels.IsChecked,
                ExportElementIds = (bool)batchExportNWC.CheckBoxExportElementIds.IsChecked,
                ExportLinks = (bool)batchExportNWC.CheckBoxExportLinks.IsChecked,
                ExportParts = (bool)batchExportNWC.CheckBoxExportParts.IsChecked,
                ExportRoomAsAttribute = (bool)batchExportNWC.CheckBoxExportRoomAsAttribute.IsChecked,
                ExportRoomGeometry = (bool)batchExportNWC.CheckBoxExportRoomGeometry.IsChecked,
                ExportUrls = (bool)batchExportNWC.CheckBoxExportUrls.IsChecked,
                FindMissingMaterials = (bool)batchExportNWC.CheckBoxFindMissingMaterials.IsChecked,
                //ConvertLights = (bool)batchExportNWC.CheckBoxConvertLights.IsChecked,
                //ConvertLinkedCADFormats = (bool)batchExportNWC.CheckBoxConvertLinkedCADFormats.IsChecked,
                //FacetingFactor = facetingFactor
            };

            switch (coordinates)
            {
                case "Общие":
                    options.Coordinates = NavisworksCoordinates.Shared;
                    break;
                case "Внутренние для проекта":
                    options.Coordinates = NavisworksCoordinates.Internal;
                    break;
            }

            switch (exportScope)
            {
                case true:
                    options.ExportScope = NavisworksExportScope.Model;
                    break;
                case false:
                    options.ExportScope = NavisworksExportScope.View;
                    options.ViewId = new FilteredElementCollector(document)
                        .OfClass(typeof(View3D))
                        .FirstOrDefault(e => e.Name == batchExportNWC.TextBoxExportScopeViewName.Text && !((View3D)e).IsTemplate)
                        .Id;
                    break;
            }

            switch (parameters)
            {
                case "Все":
                    options.Parameters = NavisworksParameters.All;
                    break;
                case "Объекты":
                    options.Parameters = NavisworksParameters.Elements;
                    break;
                case "Нет":
                    options.Parameters = NavisworksParameters.None;
                    break;
            }

            return options;
        }
        public static IFCExportOptions IFCExportOptions(Document document, IFCExportUi batchExportIFC)
        {
            IFCExportOptions options = new IFCExportOptions()
            {
                ExportBaseQuantities = (bool)batchExportIFC.CheckBoxExportBaseQuantities.IsChecked,
                FamilyMappingFile = batchExportIFC.TextBoxMapping.Text,
                FileVersion = IFCExportUi
                    .indexToIFCVersion
                    .First(e => e.Key == batchExportIFC
                        .ComboBoxIFCVersion
                        .SelectedIndex)
                    .Value,
                FilterViewId = new FilteredElementCollector(document)
                .OfClass(typeof(View))
                .FirstOrDefault(e => e.Name == batchExportIFC
                    .TextBoxExportScopeViewName
                    .Text)
                .Id,
                SpaceBoundaryLevel = batchExportIFC.ComboBoxSpaceBoundaryLevel.SelectedIndex,
                WallAndColumnSplitting = (bool)batchExportIFC.CheckBoxWallAndColumnSplitting.IsChecked
            };

            return options;
        }
        public static bool IsEverythingFilled(NWCExportUi ui)
        {
            if (ui.listBoxItems.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы один файл для экспорта!");
                return false;
            }

            string textBoxFolder = "";
            ui.Dispatcher.Invoke(() => textBoxFolder = ui.TextBoxFolder.Text);

            if (textBoxFolder == "")
            {
                MessageBox.Show("Укажите папку для экспорта!");
                return false;
            }

            if (Uri.IsWellFormedUriString(textBoxFolder, UriKind.RelativeOrAbsolute))
            {
                MessageBox.Show("Укажите корректную папку для экспорта!");
                return false;
            }

            string viewName = "";
            ui.Dispatcher.Invoke(() => viewName = ui.TextBoxExportScopeViewName.Text);

            if ((bool)ui.RadioButtonExportScopeView.IsChecked && viewName == "")
            {
                MessageBox.Show("Введите имя вида для экспорта!");
                return false;
            }

            if (!Directory.Exists(textBoxFolder))
            {
                bool isIt = true;

                MessageBoxResult messageBox = MessageBox
                    .Show("Такой папки не существует.\nСоздать папку?", "Добрый вечер", MessageBoxButton.YesNo);
                switch (messageBox)
                {
                    case MessageBoxResult.Yes:
                        Directory.CreateDirectory(textBoxFolder);
                        break;
                    case MessageBoxResult.No:
                    case MessageBoxResult.Cancel:
                        isIt = false;
                        MessageBox.Show("Нет, так нет.\nТогда живи в проклятом мире, который сам и создал.");
                        break;
                }

                if (!isIt)
                {
                    return isIt;
                }
            }

            return true;
        }
        public static bool IsEverythingFilled(IFCExportUi ui)
        {
            if (ui.listBoxItems.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы один файл для экспорта!");
                return false;
            }

            string textBoxFolder = "";
            ui.Dispatcher.Invoke(() => textBoxFolder = ui.TextBoxFolder.Text);

            if (textBoxFolder == "")
            {
                MessageBox.Show("Укажите папку для экспорта!");
                return false;
            }

            if (Uri.IsWellFormedUriString(textBoxFolder, UriKind.RelativeOrAbsolute))
            {
                MessageBox.Show("Укажите корректную папку для экспорта!");
                return false;
            }

            string viewName = "";
            ui.Dispatcher.Invoke(() => viewName = ui.TextBoxExportScopeViewName.Text);

            if ((bool)ui.RadioButtonExportScopeView.IsChecked && viewName == "")
            {
                MessageBox.Show("Введите имя вида для экспорта!");
                return false;
            }

            if (!Directory.Exists(textBoxFolder))
            {
                bool isIt = true;

                MessageBoxResult messageBox = MessageBox
                    .Show("Такой папки не существует.\nСоздать папку?", "Добрый вечер", MessageBoxButton.YesNo);
                switch (messageBox)
                {
                    case MessageBoxResult.Yes:
                        Directory.CreateDirectory(textBoxFolder);
                        break;
                    case MessageBoxResult.No:
                    case MessageBoxResult.Cancel:
                        isIt = false;
                        MessageBox.Show("Нет, так нет.\nТогда живи в проклятом мире, который сам и создал.");
                        break;
                }

                if (!isIt)
                {
                    return isIt;
                }
            }

            return true;
        }
        public static bool IsEverythingFilled(DetachModelsUi ui)
        {
            if (ui.listBoxItems.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы один файл для экспорта!");
                return false;
            }

            switch (ui.RadioButtonSavingPathMode)
            {
                case 0:
                    MessageBox.Show("Выберите режим выбора пути");
                    return false;
                case 1:
                    string textBoxFolder = "";
                    ui.Dispatcher.Invoke(() => textBoxFolder = ui.TextBoxFolder.Text);

                    if (textBoxFolder == "")
                    {
                        MessageBox.Show("Укажите папку для экспорта!");
                        return false;
                    }

                    if (Uri.IsWellFormedUriString(textBoxFolder, UriKind.RelativeOrAbsolute))
                    {
                        MessageBox.Show("Укажите корректную папку для экспорта!");
                        return false;
                    }
                    break;
                //case 2:
                //    break;
                case 3:
                    string maskIn = "";
                    string maskOut = "";
                    ui.Dispatcher.Invoke(() => maskIn = ui.TextBoxMaskIn.Text);
                    ui.Dispatcher.Invoke(() => maskOut = ui.TextBoxMaskOut.Text);

                    if (maskIn == "" || maskOut == "")
                    {
                        MessageBox.Show("Укажите маску замены");
                        return false;
                    }
                    if (!ui.listBoxItems.Select(e => e.Content)
                        .All(e => e.ToString().Contains(maskIn)))
                    {
                        MessageBox.Show("Несоответсвие входной маски и имён файлов.");
                        return false;
                    }
                    break;
            }

            return true;
        }
        public static bool IsEverythingFilled(TransmitModelsUi ui)
        {
            if (!ui.listBoxItems.Any())
            {
                MessageBox.Show("Добавьте хотя бы один файл для экспорта!");
                return false;
            }

            string textBoxFolder = "";
            ui.Dispatcher.Invoke(() => textBoxFolder = ui.TextBoxFolder.Text);

            if (textBoxFolder == "")
            {
                MessageBox.Show("Укажите папку для экспорта!");
                return false;
            }

            if (Uri.IsWellFormedUriString(textBoxFolder, UriKind.RelativeOrAbsolute))
            {
                MessageBox.Show("Укажите корректную папку для экспорта!");
                return false;
            }

            if (!Directory.Exists(textBoxFolder))
            {
                bool isIt = true;
                const string abort = "Нет, так нет.\nТогда живи в проклятом мире, который сам и создал.";

                MessageBoxResult messageBox = MessageBox
                    .Show("Такой папки не существует.\nСоздать папку?", "Добрый вечер", MessageBoxButton.YesNo);
                switch (messageBox)
                {
                    case MessageBoxResult.Yes:
                        Directory.CreateDirectory(textBoxFolder);
                        break;
                    case MessageBoxResult.No:
                        isIt = false;
                        MessageBox.Show(abort);
                        break;
                    case MessageBoxResult.Cancel:
                        isIt = false;
                        MessageBox.Show(abort);
                        break;
                }

                if (!isIt)
                {
                    return isIt;
                }
            }

            return true;
        }

        public static WorksetConfiguration CloseWorksetsWithLinks(ModelPath modelPath)
        {
            WorksetConfiguration worksetConfiguration = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);

            IList<WorksetPreview> worksets = WorksharingUtils.GetUserWorksetInfo(modelPath);
            IList<WorksetId> worksetIds = new List<WorksetId>();

            foreach (WorksetPreview worksetPreview in worksets)
            {
                if (worksetPreview.Name.StartsWith("99") || worksetPreview.Name.StartsWith("00"))
                {
                    worksetIds.Add(worksetPreview.Id);
                }
            }

            worksetConfiguration.Close(worksetIds);
            return worksetConfiguration;
        }

        public static void TaskDialogBoxShowingEvent(object sender, DialogBoxShowingEventArgs e)
        {
            TaskDialogShowingEventArgs e2 = e as TaskDialogShowingEventArgs;

            string dialogId = e2.DialogId;
            bool isConfirm = false;
            int dialogResult = 0;

            switch (dialogId)
            {
                case "TaskDialog_Missing_Third_Party_Updaters":
                    isConfirm = true;
                    dialogResult = (int)TaskDialogResult.CommandLink1;
                    break;
                case "TaskDialog_Missing_Third_Party_Updater":
                    isConfirm = true;
                    dialogResult = (int)TaskDialogResult.CommandLink1;
                    break;
                case "TaskDialog_Cannot_Find_Central_Model":
                    isConfirm = true;
                    dialogResult = (int)TaskDialogResult.Close;
                    break;
            }

            if (isConfirm)
            {
                e2.OverrideResult(dialogResult);
            }
        }

        internal static bool IsViewEmpty(Document document, Element element)
        {
            View3D view = element as View3D;

            try
            {
                using (FilteredElementCollector collector = new FilteredElementCollector(document, view.Id))
                {
                    return !collector.Where(e => e.Category != null && e.GetType() != typeof(RevitLinkInstance)).Any(e => e.CanBeHidden(view));
                }
            }
            catch
            {
                return true;
            }

        }

        public static BitmapSource GetEmbeddedImage(string name)
        {
            try
            {
                Assembly a = Assembly.GetExecutingAssembly();
                Stream s = a.GetManifestResourceStream(name);
                return BitmapFrame.Create(s);
            }
            catch
            {
                return null;
            }
        }

        public static void Application_FailuresProcessing(object sender, FailuresProcessingEventArgs e)
        {
            FailuresAccessor failuresAccessor = e.GetFailuresAccessor();
            IEnumerable<FailureMessageAccessor> failureMessages = failuresAccessor.GetFailureMessages();

            foreach (FailureMessageAccessor failureMessage in failureMessages)
            {
                if (failureMessage.GetSeverity() == FailureSeverity.Warning)
                {
                    failuresAccessor.DeleteWarning(failureMessage);
                }

                if (failureMessage.GetSeverity() == FailureSeverity.Error)
                {
                    failureMessage.SetCurrentResolutionType(FailureResolutionType.DeleteElements);
                    failuresAccessor.ResolveFailure(failureMessage);
                }
            }

            e.SetProcessingResult(FailureProcessingResult.Continue);
        }
        public static void DeleteLinks(Document document)
        {
            using (Transaction transaction = new Transaction(document))
            {
                transaction.Start("Delete links from model");

                FailureHandlingOptions failOpt = transaction.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new CopyWatchAlertSwallower());
                transaction.SetFailureHandlingOptions(failOpt);

                List<Element> links = new FilteredElementCollector(document).OfClass(typeof(RevitLinkType)).ToList();

                links.Select(link => document.Delete(link.Id));

                transaction.Commit();
            }
        }
        public static void DeleteCoordinationModels(Document document)
        {
            using (Transaction transaction = new Transaction(document))
            {
                transaction.Start("Delete navis from model");

                IList<Element> navisLinks = new FilteredElementCollector(document)
                    .OfCategory(BuiltInCategory.OST_Coordination_Model)
                    .WhereElementIsElementType()
                    .ToElements();

                navisLinks.Select(link => document.Delete(link.Id));

                transaction.Commit();
            }
        }

        public static void DetachModel(Autodesk.Revit.ApplicationServices.Application application, string filePath, DetachModelsUi ui)
        {
            ModelPath modelPath = new FilePath(filePath);

            WorksetConfiguration worksetConfiguration = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            Document document = OpenDocument.OpenDetached(application, modelPath, worksetConfiguration);

            DeleteLinks(document);

            string fileDetachedPath = "";

            switch (ui.RadioButtonSavingPathMode)
            {
                case 1:
                    string folder = "";
                    ui.Dispatcher.Invoke(() => folder = @ui.TextBoxFolder.Text);
                    fileDetachedPath = folder + "\\" + document.Title + ".rvt";
                    break;
                //case 2:
                //    string selectedItem = "";
                //    selectedItem = ui.ComboBoxPhase.SelectedItem.ToString();
                //    if (selectedItem == "ПД")
                //    {
                //        fileDetachedPath = @filePath
                //            .Replace("05_В_Работе", "06_Общие")
                //            .Replace("52_ПД", "62_ПД")
                //            .Replace(".rvt", "_отсоединено.rvt");
                //    }
                //    else
                //    {
                //        fileDetachedPath = @filePath
                //            .Replace("05_В_Работе", "06_Общие")
                //            .Replace("53_РД", "63_РД")
                //            .Replace(".rvt", "_отсоединено.rvt");
                //    }
                //    break;
                case 3:
                    string maskIn = "";
                    string maskOut = "";
                    ui.Dispatcher.Invoke(() => maskIn = @ui.TextBoxMaskIn.Text);
                    ui.Dispatcher.Invoke(() => maskOut = @ui.TextBoxMaskOut.Text);
                    fileDetachedPath = @filePath.Replace(maskIn, maskOut).Replace(".rvt", "_отсоединено.rvt");
                    break;
            }

            SaveAsOptions saveAsOptions = new SaveAsOptions()
            {
                OverwriteExistingFile = true,
                MaximumBackups = 1
            };
            WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions()
            {
                SaveAsCentral = true
            };
            saveAsOptions.SetWorksharingOptions(worksharingSaveAsOptions);

            ModelPath modelDetachedPath = new FilePath(fileDetachedPath);
            document.SaveAs(modelDetachedPath, saveAsOptions);

            try
            {
                FreeTheModel(document);
            }
            catch
            {
            }

            document.Close();
            document.Dispose();
        }

        public static void UnloadRevitLinks(ModelPath location)
        ///  This method will set all Revit links to be unloaded the next time the document at the given location is opened. 
        ///  The TransmissionData for a given document only contains top-level Revit links, not nested links.
        ///  However, nested links will be unloaded if their parent links are unloaded, so this function only needs to look at the document's immediate links. 
        {
            // access transmission data in the given Revit file
            TransmissionData transData = TransmissionData.ReadTransmissionData(location);

            if (transData != null)
            {
                // collect all (immediate) external references in the model
                ICollection<ElementId> externalReferences = transData.GetAllExternalFileReferenceIds();

                // find every reference that is a link
                foreach (ElementId refId in externalReferences)
                {
                    ExternalFileReference extRef = transData.GetLastSavedReferenceData(refId);
                    if (extRef.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink)
                    {
                        //ModelPath modelPath = extRef.GetAbsolutePath();
                        //string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                        // we do not want to change neither the path nor the path-type
                        // we only want the links to be unloaded (shouldLoad = false)
                        try
                        {
                            transData.SetDesiredReferenceData(refId, extRef.GetPath(), extRef.PathType, false);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                // make sure the IsTransmitted property is set 
                transData.IsTransmitted = true;

                // modified transmission data must be saved back to the model
                TransmissionData.WriteTransmissionData(location, transData);
            }
            else
            {
                TaskDialog.Show("Unload Links", "The document does not have any transmission data");
            }
        }

        public static void ReplaceRevitLinks(ModelPath filePath, Dictionary<string, string> oldNewFilePairs)
        {
            TransmissionData transData = TransmissionData.ReadTransmissionData(filePath);

            if (transData != null)
            {
                ICollection<ElementId> externalReferences = transData.GetAllExternalFileReferenceIds();

                foreach (ElementId refId in externalReferences)
                {
                    ExternalFileReference extRef = transData.GetLastSavedReferenceData(refId);
                    ModelPath modelPath = extRef.GetAbsolutePath();
                    string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                    if (extRef.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink && oldNewFilePairs.Any(e => e.Key == path))
                    {
                        string newFile = oldNewFilePairs.FirstOrDefault(e => e.Key == path).Value;
                        ModelPath newPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(newFile);
                        try
                        {
                            transData.SetDesiredReferenceData(refId, newPath, PathType.Absolute, true);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                            continue;
                        }
                    }
                }

                transData.IsTransmitted = true;

                TransmissionData.WriteTransmissionData(filePath, transData);
            }
            else
            {
                TaskDialog.Show("Replace Links", "The document does not have any transmission data");
            }
        }

        public static void BatchExportNWC(UIApplication uiApp, NWCExportUi ui, ref Logger logger)
        {
            Autodesk.Revit.ApplicationServices.Application application = uiApp.Application;

            List<ListBoxItem> listItems = ui.listBoxItems.ToList();

            foreach (ListBoxItem item in listItems)
            {
                string filePath = item.Content.ToString();

                bool fileIsWorkshared = true;

                logger.LineBreak();
                DateTime startTime = DateTime.Now;
                logger.Start(filePath);

                if (!File.Exists(filePath))
                {
                    logger.Error($"Файла {filePath} не существует. Ты совсем Туттуру?");
                    item.Background = Brushes.Red;
                    continue;
                }
                uiApp.DialogBoxShowing += new EventHandler<DialogBoxShowingEventArgs>(TaskDialogBoxShowingEvent);
                application.FailuresProcessing += new EventHandler<FailuresProcessingEventArgs>(Application_FailuresProcessing);
                Document document;

                BasicFileInfo fileInfo;

                try
                {
                    fileInfo = BasicFileInfo.Extract(filePath);
                    if (!fileInfo.IsWorkshared)
                    {
                        document = application.OpenDocumentFile(filePath);
                        fileIsWorkshared = false;
                    }
                    else if (filePath.Equals(fileInfo.CentralPath))
                    {
                        ModelPath modelPath = new FilePath(filePath);
                        WorksetConfiguration worksetConfiguration = CloseWorksetsWithLinks(modelPath);
                        document = OpenDocument.OpenAsIs(application, modelPath, worksetConfiguration);
                    }
                    else
                    {
                        ModelPath modelPath = new FilePath(filePath);
                        WorksetConfiguration worksetConfiguration = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
                        document = OpenDocument.OpenAsIs(application, modelPath, worksetConfiguration);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error("Файл не открылся. ", ex);
                    item.Background = Brushes.Red;
                    continue;
                }

                fileInfo.Dispose();
                logger.FileOpened();

                item.Background = Brushes.Blue;
                bool isFuckedUp = false;

                try
                {
                    ExportModelToNWC(document, ui, ref isFuckedUp, logger);
                }
                catch (Exception ex)
                {
                    logger.Error("Ля, я хз даже. Смотри, что в исключении написано: ", ex);
                    isFuckedUp = true;
                }
                finally
                {
                    if (fileIsWorkshared)
                    {
                        try
                        {
                            FreeTheModel(document);
                        }
                        catch (Exception ex)
                        {
                            logger.Error("Не смог освободить рабочие наборы. ", ex);
                            isFuckedUp = true;
                        }
                    }

                    document.Close(false);
                    document.Dispose();

                    if (isFuckedUp)
                    {
                        item.Background = Brushes.Red;
                    }
                    else
                    {
                        item.Background = Brushes.Green;
                        logger.Success("Всё ок.");
                    }

                    uiApp.DialogBoxShowing -= new EventHandler<DialogBoxShowingEventArgs>(TaskDialogBoxShowingEvent);
                    application.FailuresProcessing -= new EventHandler<FailuresProcessingEventArgs>(Application_FailuresProcessing);
                    logger.TimeForFile(startTime);
                    Thread.Sleep(500);
                }
            }

            application.Dispose();

            logger.LineBreak();
            logger.ErrorTotal();
            logger.TimeTotal();
            logger.Dispose();
        }
        public static void BatchExportIFC(UIApplication uiApp, IFCExportUi ui, ref Logger logger)
        {
            Autodesk.Revit.ApplicationServices.Application application = uiApp.Application;

            List<ListBoxItem> listItems = ui.listBoxItems.ToList();

            foreach (ListBoxItem item in listItems)
            {
                string filePath = item.Content.ToString();

                bool fileIsWorkshared = true;

                logger.LineBreak();
                DateTime startTime = DateTime.Now;
                logger.Start(filePath);

                if (!File.Exists(filePath))
                {
                    logger.Error($"Файла {filePath} не существует. Ты совсем Туттуру?");
                    item.Background = Brushes.Red;
                    continue;
                }
                uiApp.DialogBoxShowing += new EventHandler<DialogBoxShowingEventArgs>(TaskDialogBoxShowingEvent);
                application.FailuresProcessing += new EventHandler<FailuresProcessingEventArgs>(Application_FailuresProcessing);
                Document document;

                BasicFileInfo fileInfo;

                try
                {
                    fileInfo = BasicFileInfo.Extract(filePath);
                    if (!fileInfo.IsWorkshared)
                    {
                        document = application.OpenDocumentFile(filePath);
                        fileIsWorkshared = false;
                    }
                    else if (filePath.Equals(fileInfo.CentralPath))
                    {
                        ModelPath modelPath = new FilePath(filePath);
                        WorksetConfiguration worksetConfiguration = CloseWorksetsWithLinks(modelPath);
                        document = OpenDocument.OpenAsIs(application, modelPath, worksetConfiguration);
                    }
                    else
                    {
                        ModelPath modelPath = new FilePath(filePath);
                        WorksetConfiguration worksetConfiguration = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
                        document = OpenDocument.OpenAsIs(application, modelPath, worksetConfiguration);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error("Файл не открылся. ", ex);
                    item.Background = Brushes.Red;
                    continue;
                }

                fileInfo.Dispose();
                logger.FileOpened();

                item.Background = Brushes.Blue;
                bool isFuckedUp = false;

                try
                {
                    ExportModelToIFC(document, ui, ref isFuckedUp, logger);
                }
                catch (Exception ex)
                {
                    logger.Error("Ля, я хз даже. Смотри, что в исключении написано: ", ex);
                    isFuckedUp = true;
                }
                finally
                {
                    if (fileIsWorkshared)
                    {
                        try
                        {
                            FreeTheModel(document);
                        }
                        catch (Exception ex)
                        {
                            logger.Error("Не смог освободить рабочие наборы. ", ex);
                            isFuckedUp = true;
                        }
                    }

                    document.Close(false);
                    document.Dispose();

                    if (isFuckedUp)
                    {
                        item.Background = Brushes.Red;
                    }
                    else
                    {
                        item.Background = Brushes.Green;
                        logger.Success("Всё ок.");
                    }

                    uiApp.DialogBoxShowing -= new EventHandler<DialogBoxShowingEventArgs>(TaskDialogBoxShowingEvent);
                    application.FailuresProcessing -= new EventHandler<FailuresProcessingEventArgs>(Application_FailuresProcessing);
                    logger.TimeForFile(startTime);
                    Thread.Sleep(500);
                }
            }

            application.Dispose();

            logger.LineBreak();
            logger.ErrorTotal();
            logger.TimeTotal();
            logger.Dispose();
        }

        public static string MD5Hash(string fileName)
        {
            using (MD5 md5 = MD5.Create())
            {
                try
                {
                    using (FileStream stream = File.OpenRead(fileName))
                    {
                        return Convert.ToBase64String(md5.ComputeHash(stream));
                    }
                }
                catch
                {
                    return null;
                }
            }
        }

        public static void FreeTheModel(Document document)
        {
            RelinquishOptions relinquishOptions = new RelinquishOptions(true);
            TransactWithCentralOptions transactWithCentralOptions = new TransactWithCentralOptions();
            WorksharingUtils.RelinquishOwnership(document, relinquishOptions, transactWithCentralOptions);
        }
    }
}
