﻿using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Appolo.BatchExport.Utils
{
    public class CopyWatchAlertSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
        {
            IList<FailureMessageAccessor> failures = a.GetFailureMessages();

            foreach (FailureMessageAccessor f in failures)
            {
                FailureDefinitionId id = f.GetFailureDefinitionId();

                if (BuiltInFailures.CopyMonitorFailures.CopyWatchAlert == id)
                {
                    a.DeleteWarning(f);
                }
            }

            return FailureProcessingResult.Continue;
        }
    }
}
