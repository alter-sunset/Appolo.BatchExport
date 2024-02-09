using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Appolo.BatchExport.IFC
{
    public class IFCForm : IFCExportOptions
    {
        public string DestinationFolder { get; set; }
        public string Prefix { get; set; }
        public string Postfix { get; set; }
        public List<string> RVTFiles { get; set; }
        public string ViewName { get; set; }
        public bool ExportView { get; set; }
    }
}
