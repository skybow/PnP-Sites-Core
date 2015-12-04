using System.Collections.Generic;
using OfficeDevPnP.Core.Entities;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class ListPage
    {
        public string PageUrl { get; set; }

        public List<WebPartEntity> WebPartEntities { get; set; } 
    }
}
