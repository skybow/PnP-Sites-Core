namespace Microsoft.SharePoint.Client
{
    public static class ListItemExtensions
    {
        public static bool IsFile(this ListItem listItem)
        {
            return listItem.FileSystemObjectType == FileSystemObjectType.File;
        }
    }
}
