using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts
{
    interface IWebPartTokenizer
    {
        string Tokenize(string xml, TokenParser parser);
    }
}
