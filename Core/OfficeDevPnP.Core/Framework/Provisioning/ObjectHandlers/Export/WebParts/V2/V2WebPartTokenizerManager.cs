using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts.V2
{
    class V2WebPartTokenizerManager
    {
        public static IWebPartTokenizer GetWebPartTokenizer(string webPartType) {
            switch (webPartType)
            {
                default:
                    return new V2DefaultWebPartTokenizer();
                    break;
            }
        }
    }
}
