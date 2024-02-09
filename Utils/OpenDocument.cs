using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;

namespace Appolo.BatchExport
{
    public static class OpenDocument
    {
        public static Document OpenAsIs(Application application, ModelPath modelPath, WorksetConfiguration worksetConfiguration)
        {
            OpenOptions openOptions = new OpenOptions()
            {
                DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
            };

            openOptions.SetOpenWorksetsConfiguration(worksetConfiguration);
            Document openedDoc = application.OpenDocumentFile(modelPath, openOptions);

            return openedDoc;
        }

        public static Document OpenDetached(Application application, ModelPath modelPath, WorksetConfiguration worksetConfiguration)
        {
            OpenOptions openOptions = new OpenOptions()
            {
                DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            };

            openOptions.SetOpenWorksetsConfiguration(worksetConfiguration);
            Document openedDoc = application.OpenDocumentFile(modelPath, openOptions);

            return openedDoc;
        }

        public static Document OpenTransmitted(Application application, ModelPath modelPath)
        {
            OpenOptions openOptions = new OpenOptions()
            {
                DetachFromCentralOption = DetachFromCentralOption.ClearTransmittedSaveAsNewCentral
            };

            WorksetConfiguration worksetConfiguration = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            openOptions.SetOpenWorksetsConfiguration(worksetConfiguration);

            Document openedDoc = application.OpenDocumentFile(modelPath, openOptions);

            return openedDoc;

        }
    }
}
