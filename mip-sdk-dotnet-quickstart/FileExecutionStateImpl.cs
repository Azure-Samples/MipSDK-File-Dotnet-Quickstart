using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.Policy;

namespace MipSdkDotNetQuickstart
{
    class FileExecutionStateImpl : IFileExecutionState
    {
        public byte[] SerializedProtectionInfo
        {
            get { return new byte[2]; }
            set { } 
        }

        public Dictionary<string, ClassificationResult> GetClassificationResults(IFileHandler fileHandler, List<ClassificationRequest> classificationIds)
        {
            Dictionary<string, ClassificationResult> result = new Dictionary<string, ClassificationResult>();

            //simulate classification and result. 
            result.Add("0", new ClassificationResult() { ConfidenceLevel = 76, Count = 5, Id = "MyId" });

            return result;
        }
    }
}
