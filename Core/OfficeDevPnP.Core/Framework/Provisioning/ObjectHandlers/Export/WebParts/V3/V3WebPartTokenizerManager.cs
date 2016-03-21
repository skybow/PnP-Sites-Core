using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts.V3
{
    class V3WebPartTokenizerManager
    {
        public static IWebPartTokenizer GetWebPartTokenizer(string webPartType)
        {
            switch (webPartType)
            {
                default:
                    return new V3DefaultWebPartTokenizer();
                    break;
            }
        }
    }
}
