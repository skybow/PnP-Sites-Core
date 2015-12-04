using System;
using System.Linq;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class Form
    {
        public PageType FormType { get; set; }
        public string ServerRelativeUrl { get; set; }
        public bool IsDefault { get; set; }

        //public string GetFormName()
        //{
        //    return String.IsNullOrEmpty(ServerRelativeUrl) ? string.Empty : this.ServerRelativeUrl.Split('/').Last();
        //}

        internal static string GetNameByType(PageType formType)
        {
            switch (formType)
            {
                case PageType.DisplayForm:
                    return "DISPLAY";
                case PageType.EditForm:
                    return "EDIT";
                case PageType.NewForm:
                    return "NEW";
                default:
                    return null;
            }
        }
        internal static PageType GetTypeByName(string name)
        {
            switch (name)
            {
                case "DISPLAY":
                    return PageType.DisplayForm;
                case "EDIT":
                    return PageType.EditForm;
                case "NEW":
                    return PageType.NewForm;
                default:
                    return PageType.Invalid;
            }
        }
    }
}
