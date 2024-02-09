using Appolo.BatchExport.NWC;
using Appolo.Utilities;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace Appolo.BatchExport
{
    /// <summary>
    /// This is an example of of wrapping a method with an ExternalEventHandler using a string argument.
    /// Any type of argument can be passed to the RevitEventWrapper, and therefore be used in the execution
    /// of a method which has to take place within a "Valid Revit API Context".
    /// </summary>
    public class EventHandlerWithStringArg : RevitEventWrapper<string>
    {
        /// <summary>
        /// The Execute override void must be present in all methods wrapped by the RevitEventWrapper.
        /// This defines what the method will do when raised externally.
        /// </summary>
        public override void Execute(UIApplication uiApp, string args)
        {
            // Do your processing here with "args"
            TaskDialog.Show("External Event", args);
        }
    }

    public class EventHandlerNWCExportBatchUiArg : RevitEventWrapper<NWCExportUi>
    {
        public override void Execute(UIApplication uiApp, NWCExportUi ui)
        {
            if (ui.ListBoxJsonConfigs.Items.Count == 0)
            {
                MessageBox.Show("Загрузите конфиги.");
                return;
            }

            DateTime timeStart = DateTime.Now;

            foreach (string config in ui.ListBoxJsonConfigs.Items)
            {
                try
                {
                    using (StreamReader file = File.OpenText(config))
                    {
                        JsonSerializer serializer = new JsonSerializer();

                        NWCForm form = (NWCForm)serializer.Deserialize(file, typeof(NWCForm));
                        ui.NWCFormDeserilaizer(form);
                    }
                }
                catch
                {
                    continue;
                }

                string folder = "";
                ui.Dispatcher.Invoke(() => folder = @ui.TextBoxFolder.Text);
                Logger logger = new Logger(folder);

                Methods.BatchExportNWC(uiApp, ui, ref logger);
                Thread.Sleep(3000);
            }

            TaskDialog taskDialog = new TaskDialog("Готово!")
            {
                CommonButtons = TaskDialogCommonButtons.Close,
                Id = "ExportBatchNWCFinished",
                MainContent = $"Задание выполнено. Общее время выполнения: {DateTime.Now - timeStart}"
            };

            ui.IsEnabled = false;
            taskDialog.Show();
            ui.IsEnabled = true;
        }
    }

    public class EventHandlerNWCExportUiArg : RevitEventWrapper<NWCExportUi>
    {
        public override void Execute(UIApplication uiApp, NWCExportUi ui)
        {
            if (!Methods.IsEverythingFilled(ui))
            {
                return;
            }

            string folder = "";
            ui.Dispatcher.Invoke(() => folder = @ui.TextBoxFolder.Text);
            Logger logger = new Logger(folder);

            Methods.BatchExportNWC(uiApp, ui, ref logger);

            TaskDialog taskDialog = new TaskDialog("Готово!")
            {
                CommonButtons = TaskDialogCommonButtons.Close,
                Id = "ExportNWCFinished",
                MainContent = $"В процессе выполнения было {logger.ErrorCount} ошибок из {logger.ErrorCount + logger.SuccessCount} файлов."
            };

            ui.IsEnabled = false;
            taskDialog.Show();
            ui.IsEnabled = true;
        }
    }
    public class EventHandlerIFCExportUiArg : RevitEventWrapper<IFCExportUi>
    {
        public override void Execute(UIApplication uiApp, IFCExportUi ui)
        {
            if (!Methods.IsEverythingFilled(ui))
            {
                return;
            }

            string folder = "";
            ui.Dispatcher.Invoke(() => folder = @ui.TextBoxFolder.Text);
            Logger logger = new Logger(folder);

            Methods.BatchExportIFC(uiApp, ui, ref logger);

            TaskDialog taskDialog = new TaskDialog("Готово!")
            {
                CommonButtons = TaskDialogCommonButtons.Close,
                Id = "ExportIFCFinished",
                MainContent = $"В процессе выполнения было {logger.ErrorCount} ошибок из {logger.ErrorCount + logger.SuccessCount} файлов."
            };

            ui.IsEnabled = false;
            taskDialog.Show();
            ui.IsEnabled = true;
        }
    }

    public class EventHandlerDetachModelsUiArg : RevitEventWrapper<DetachModelsUi>
    {
        public override void Execute(UIApplication uiApp, DetachModelsUi ui)
        {
            if (!Methods.IsEverythingFilled(ui))
            {
                return;
            }

            Application application = uiApp.Application;
            List<ListBoxItem> listItems = @ui.listBoxItems.ToList();

            foreach (ListBoxItem item in listItems)
            {
                string filePath = item.Content.ToString();

                if (!File.Exists(filePath))
                {
                    string error = $"Файла {filePath} не существует. Ты совсем Туттуру?";
                    item.Background = Brushes.Red;
                    continue;
                }

                uiApp.DialogBoxShowing += new EventHandler<DialogBoxShowingEventArgs>(Methods.TaskDialogBoxShowingEvent);
                application.FailuresProcessing += new EventHandler<Autodesk.Revit.DB.Events.FailuresProcessingEventArgs>(Methods.Application_FailuresProcessing);

                Methods.DetachModel(application, filePath, ui);

                uiApp.DialogBoxShowing -= new EventHandler<DialogBoxShowingEventArgs>(Methods.TaskDialogBoxShowingEvent);
                application.FailuresProcessing -= new EventHandler<Autodesk.Revit.DB.Events.FailuresProcessingEventArgs>(Methods.Application_FailuresProcessing);
            }

            TaskDialog taskDialog = new TaskDialog("Готово!")
            {
                CommonButtons = TaskDialogCommonButtons.Close,
                Id = "DetachModelsFinished",
                MainContent = "Задание выполнено"
            };

            application.Dispose();

            ui.IsEnabled = false;
            taskDialog.Show();
            ui.IsEnabled = true;
        }
    }

    public class EventHandlerTransmitModelsUiArg : RevitEventWrapper<TransmitModelsUi>
    {
        public override void Execute(UIApplication uiApp, TransmitModelsUi ui)
        {
            if (!Methods.IsEverythingFilled(ui))
            {
                return;
            }

            Application application = uiApp.Application;
            List<ListBoxItem> listItems = @ui.listBoxItems.ToList();

            foreach (ListBoxItem item in listItems)
            {
                string filePath = item.Content.ToString();

                if (!File.Exists(filePath))
                {
                    string error = $"Файла {filePath} не существует. Ты совсем Туттуру?";
                    item.Background = Brushes.Red;
                    continue;
                }

                string folder = "";
                ui.Dispatcher.Invoke(() => folder = @ui.TextBoxFolder.Text);
                string transmittedFilePath = folder + "\\" + filePath.Split('\\').Last();

                File.Copy(filePath, transmittedFilePath, true);

                ModelPath transmittedModelPath = new FilePath(transmittedFilePath);
                Methods.UnloadRevitLinks(transmittedModelPath);
            }

            TaskDialog taskDialog = new TaskDialog("Готово!")
            {
                CommonButtons = TaskDialogCommonButtons.Close,
                Id = "TransmitModelsFinished",
                MainContent = "Задание выполнено"
            };

            application.Dispose();

            ui.IsEnabled = false;
            taskDialog.Show();
            ui.IsEnabled = true;
        }
    }

    public class EventHandlerMigrateModelsUiArg : RevitEventWrapper<MigrateModelsUi>
    {
        public override void Execute(UIApplication uiApp, MigrateModelsUi ui)
        {
            if (string.IsNullOrEmpty(ui.TextBoxConfig.Text)
                || !ui.TextBoxConfig.Text.EndsWith(".json", true, CultureInfo.InvariantCulture))
            {
                MessageBox.Show("Предоставьте ссылку на конфиг");
                return;
            }

            Dictionary<string, string> items;
            List<string> movedFiles = new List<string>();
            List<string> failedFiles = new List<string>();

            using (StreamReader file = File.OpenText(ui.TextBoxConfig.Text))
            {
                try
                {
                    JsonSerializer serializer = new JsonSerializer();
                    items = (Dictionary<string, string>)serializer.Deserialize(file, typeof(Dictionary<string, string>));
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("Неверная схема файла");
                    return;
                }
            }

            Application application = uiApp.Application;

            foreach (string oldFile in items.Keys)
            {
                if (!File.Exists(oldFile))
                {
                    failedFiles.Add(oldFile);
                    continue;
                }

                string newFile = items.FirstOrDefault(e => e.Key == oldFile).Value;

                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(newFile));
                    File.Copy(oldFile, newFile, true);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                    failedFiles.Add(oldFile);
                    continue;
                }

                movedFiles.Add(newFile);
            }

            foreach (string newFile in movedFiles)
            {
                ModelPath newFilePath = new FilePath(newFile);
                Methods.ReplaceRevitLinks(newFilePath, items);

                Document document = OpenDocument.OpenTransmitted(application, newFilePath);

                newFilePath.Dispose();

                try
                {
                    Methods.FreeTheModel(document);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                document.Close();
                document.Dispose();
            }

            TaskDialog taskDialog = new TaskDialog("Готово!")
            {
                CommonButtons = TaskDialogCommonButtons.Close,
                Id = "MigrateModelsFinished",
                MainContent = $"Задание выполнено.\nСледующие файлы не были скопированы:\n{string.Join("\n", failedFiles)}"
            };

            application.Dispose();

            ui.IsEnabled = false;
            taskDialog.Show();
            ui.IsEnabled = true;
        }
    }
}
