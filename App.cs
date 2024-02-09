using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Media.Imaging;
using System.Windows;

namespace Appolo.BatchExport
{
    /// <summary>
    /// This is the main class which defines the Application, and inherits from Revit's
    /// IExternalApplication class.
    /// </summary>
    public class App : IExternalApplication
    {
        // class instance
        public static App ThisApp;

        // ModelessForm instance
        public static NWCExportUi _mMyFormNWC;
        public static IFCExportUi _mMyFormIFC;
        public static DetachModelsUi _mMyFormDetach;
        public static TransmitModelsUi _mMyFormTransmit;
        public static MigrateModelsUi _mMyFormMigrate;

        public Result OnStartup(UIControlledApplication a)
        {
            _mMyFormNWC = null; // no dialog needed yet; the command will bring it
            _mMyFormIFC = null;
            _mMyFormDetach = null;
            ThisApp = this; // static access to this application instance

            // Method to add Tab and Panel 
            RibbonPanel panel = RibbonPanel(a);
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            string[] applIconPath = { "Appolo.BatchExport.Appolo.png", "Appolo.BatchExport.Appolo_16.png" };
            string[] ifcIconPath = { "Appolo.BatchExport.ifc.png", "Appolo.BatchExport.ifc_16.png" };
            string[] navisIconPath = { "Appolo.BatchExport.navisworks.png", "Appolo.BatchExport.navisworks_16.png" };

            PushButtonData exportModelsToNWCButtonData = new PushButtonData(
                   "Экспорт NWC",
                   "Экспорт\nNWC",
                   thisAssemblyPath,
                   "Appolo.BatchExport.ExportModelsToNWC")
            {
                AvailabilityClassName = "Appolo.BatchExport.ExportModelsToNWCCommand_Availability"
            };

            string exportModelsToNWCToolTip = "Пакетный экспорт в NWC";
            CreateNewPushButton(panel, exportModelsToNWCButtonData, exportModelsToNWCToolTip, navisIconPath);

            PushButtonData exportModelsDetachedButtonData = new PushButtonData(
                   "Экспорт отсоединённых моделей",
                   "Экспорт\nотсоединённых\nмоделей",
                   thisAssemblyPath,
                   "Appolo.BatchExport.ExportModelsDetached")
            {
                AvailabilityClassName = "Appolo.BatchExport.ExportModelsDetachedCommand_Availability"
            };

            string exportModelsDetachedToolTip = "Пакетный экспорт отсоединённых моделей";
            CreateNewPushButton(panel, exportModelsDetachedButtonData, exportModelsDetachedToolTip, applIconPath);

            PushButtonData transmitModelsButtonData = new PushButtonData(
                   "Передача моделей",
                   "Передача",
                   thisAssemblyPath,
                   "Appolo.BatchExport.ExportModelsTransmitted")
            {
                AvailabilityClassName = "Appolo.BatchExport.ExportModelsTransmittedCommand_Availability"
            };

            string transmitModelsToolTip = "Пакетная передача моделей";
            CreateNewPushButton(panel, transmitModelsButtonData, transmitModelsToolTip, applIconPath);

            PushButtonData migrateModelsButtonData = new PushButtonData(
                   "Миграция моделей",
                   "Миграция\nмоделей",
                   thisAssemblyPath,
                   "Appolo.BatchExport.MigrateModels")
            {
                AvailabilityClassName = "Appolo.BatchExport.MigrateModelsCommand_Availability"
            };

            string migrateModelsToolTip = "Пакетная миграция моделей с обновлением связей";
            CreateNewPushButton(panel, migrateModelsButtonData, migrateModelsToolTip, applIconPath);

            PushButtonData exportModelsToIFCButtonData = new PushButtonData(
                   "Экспорт IFC",
                   "Экспорт\nIFC",
                   thisAssemblyPath,
                   "Appolo.BatchExport.ExportModelsToIFC")
            {
                AvailabilityClassName = "Appolo.BatchExport.ExportModelsToIFCCommand_Availability"
            };

            string exportModelsToIFCToolTip = "Пакетный экспорт в IFC";
            CreateNewPushButton(panel, exportModelsToIFCButtonData, exportModelsToIFCToolTip, ifcIconPath);

            return Result.Succeeded;
        }
        static void CreateNewPushButton(RibbonPanel ribbonPanel, PushButtonData pushButtonData, string toolTip, string[] iconPath)
        {
            BitmapSource bitmap_32 = Methods.GetEmbeddedImage(iconPath[0]);
            BitmapSource bitmap_16 = Methods.GetEmbeddedImage(iconPath[1]);
            PushButton pushButton = ribbonPanel.AddItem(pushButtonData) as PushButton;
            pushButton.ToolTip = toolTip;
            pushButton.Image = bitmap_16;
            pushButton.LargeImage = bitmap_32;
        }

        /// <summary>
        /// What to do when the application is shut down.
        /// </summary>
        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }

        /// <summary>
        /// This is the method which launches the WPF window, and injects any methods that are
        /// wrapped by ExternalEventHandlers. This can be done in a number of different ways, and
        /// implementation will differ based on how the WPF is set up.
        /// </summary>
        /// <param name="uiapp">The Revit UIApplication within the add-in will operate.</param>
        public void ShowFormNWC(UIApplication uiapp)
        {
            // If we do not have a dialog yet, create and show it
            if (_mMyFormNWC != null && _mMyFormNWC == null) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerNWCExportUiArg evUi = new EventHandlerNWCExportUiArg();
            EventHandlerNWCExportBatchUiArg eventHandlerNWCExportBatchUiArg = new EventHandlerNWCExportBatchUiArg();

            // The dialog becomes the owner responsible for disposing the objects given to it.
            _mMyFormNWC = new NWCExportUi(uiapp, evUi, eventHandlerNWCExportBatchUiArg) { Height = 900, Width = 800 };
            _mMyFormNWC.Show();
        }
        public void ShowFormIFC(UIApplication uiapp)
        {
            // If we do not have a dialog yet, create and show it
            if (_mMyFormIFC != null && _mMyFormIFC == null) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerIFCExportUiArg evUi = new EventHandlerIFCExportUiArg();

            // The dialog becomes the owner responsible for disposing the objects given to it.
            _mMyFormIFC = new IFCExportUi(uiapp, evUi) { Height = 700, Width = 800 };
            _mMyFormIFC.Show();
        }
        public void ShowFormDetach(UIApplication uiapp)
        {
            // If we do not have a dialog yet, create and show it
            if (_mMyFormDetach != null && _mMyFormDetach == null) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerDetachModelsUiArg evUi = new EventHandlerDetachModelsUiArg();

            // The dialog becomes the owner responsible for disposing the objects given to it.

            try
            {
                _mMyFormDetach = new DetachModelsUi(uiapp, evUi) { Height = 600, Width = 800 };
                _mMyFormDetach.Show();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void ShowFormTransmit(UIApplication uiapp)
        {
            // If we do not have a dialog yet, create and show it
            if (_mMyFormTransmit != null && _mMyFormTransmit == null) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerTransmitModelsUiArg evUi = new EventHandlerTransmitModelsUiArg();

            // The dialog becomes the owner responsible for disposing the objects given to it.

            try
            {
                _mMyFormTransmit = new TransmitModelsUi(uiapp, evUi) { Height = 500, Width = 800 };
                _mMyFormTransmit.Show();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void ShowFormMigrate(UIApplication uiapp)
        {
            // If we do not have a dialog yet, create and show it
            if (_mMyFormMigrate != null && _mMyFormMigrate == null) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerMigrateModelsUiArg evUi = new EventHandlerMigrateModelsUiArg();

            // The dialog becomes the owner responsible for disposing the objects given to it.

            try
            {
                _mMyFormMigrate = new MigrateModelsUi(uiapp, evUi) { Height = 200, Width = 600 };
                _mMyFormMigrate.Show();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #region Ribbon Panel

        public RibbonPanel RibbonPanel(UIControlledApplication a)
        {
            string tab = "Appolo"; // Tab name
            // Empty ribbon panel 
            RibbonPanel ribbonPanel = null;
            // Try to create ribbon tab. 
            try
            {
                a.CreateRibbonTab(tab);
            }
            catch (Exception ex)
            {
                Util.HandleError(ex);
            }

            // Try to create ribbon panel.
            try
            {
                RibbonPanel panel = a.CreateRibbonPanel(tab, "Пакетный экспорт");
            }
            catch (Exception ex)
            {
                Util.HandleError(ex);
            }

            // Search existing tab for your panel.
            List<RibbonPanel> panels = a.GetRibbonPanels(tab);
            foreach (RibbonPanel p in panels.Where(p => p.Name == "Пакетный экспорт"))
            {
                ribbonPanel = p;
            }

            //return panel 
            return ribbonPanel;
        }

        #endregion
    }
}
