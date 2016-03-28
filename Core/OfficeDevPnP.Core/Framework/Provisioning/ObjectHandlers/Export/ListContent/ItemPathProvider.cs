using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using List = Microsoft.SharePoint.Client.List;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.ListContent
{
    public class ItemPathProvider
    {
        public const string FIELD_ItemDir = "FileDirRef";
        public const string FIELD_ItemName = "FileLeafRef";
        public const string FIELD_ItemType = "FSObjType";

        public const string FIELD_ItemType_FolderValue = "1";
        
        private string m_listServerRelativeUrl = null;
        private List m_list = null;
        private Web m_web = null;

        public ItemPathProvider(List list, Web web)
        {
            var fields = list.Fields;

            m_listServerRelativeUrl = list.RootFolder.ServerRelativeUrl;
            m_list = list;
            m_web = web;
        }
        
        public ListItemCreationInformation GetItemCreationInformation(DataRow dataRow)
        {
            ListItemCreationInformation creationInfo = null;

            string dir = null;
            string dirValue;
            if (dataRow.Values.TryGetValue(FIELD_ItemDir, out dirValue) && !string.IsNullOrEmpty(dirValue))
            {
                dir = TokenParser.CombineUrl(this.m_web.ServerRelativeUrl, dirValue);                
            }

            string objType;
            if (dataRow.Values.TryGetValue(FIELD_ItemType, out objType))
            {
                if (objType == FIELD_ItemType_FolderValue)
                {
                    string name;
                    if (dataRow.Values.TryGetValue(FIELD_ItemName, out name))
                    {
                        creationInfo = new ListItemCreationInformation()
                        {
                            UnderlyingObjectType = FileSystemObjectType.Folder,
                            LeafName = name,
                            FolderUrl = dir
                        };
                    }
                }
            }

            if( null == creationInfo )
            {
                creationInfo = new ListItemCreationInformation()
                {
                    FolderUrl = dir
                };
            }

            return creationInfo;
        }

        public void ExtractItemPathValues(ListItem item, Dictionary<string,string> dataRowValues)
        {
            string dir = item[FIELD_ItemDir] as string; ;
            if (!string.IsNullOrEmpty(dir) && 
                !dir.Equals(m_listServerRelativeUrl, StringComparison.OrdinalIgnoreCase) &&
                dir.StartsWith(m_web.ServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                dataRowValues[FIELD_ItemDir] = dir.Substring(m_web.ServerRelativeUrl.Length);
            }

            string sObjType = item[FIELD_ItemType] as string;
            if (sObjType == FIELD_ItemType_FolderValue) //Folder
            {
                dataRowValues[FIELD_ItemType] = FIELD_ItemType_FolderValue;                
                
                string name = item[FIELD_ItemName] as string;
                if (!string.IsNullOrEmpty(name))
                {
                    dataRowValues[FIELD_ItemName] = name;
                }
            }
        }

        /*
        private Folder EnsureFolder( string urlWebRelative  )
        {
            Folder currentFolder = m_list.RootFolder;
            string rootUrl = currentFolder.ServerRelativeUrl;            

            string targetFolderServerUrl = TokenParser.CombineUrl( this.m_web.ServerRelativeUrl, urlWebRelative );

            // Get remaining parts of the path and split

            var folderRootRelativeUrl = targetFolderServerUrl.Substring(currentFolder.ServerRelativeUrl.Length);
            var childFolderNames = folderRootRelativeUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var currentCount = 0;

            foreach (var folderName in childFolderNames)
            {
                currentCount++;

                // Find next part of the path
                var folderCollection = currentFolder.Folders;
                folderCollection.Context.Load(folderCollection);
                folderCollection.Context.ExecuteQueryRetry();
                Folder nextFolder = null;
                foreach (Folder existingFolder in folderCollection)
                {
                    if (string.Equals(existingFolder.Name, folderName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        nextFolder = existingFolder;
                        break;
                    }
                }

                // Or create it
                if (nextFolder == null)
                {
                    ListItem itemFolder = this.m_list.AddItem(new ListItemCreationInformation()
                    {
                        UnderlyingObjectType = FileSystemObjectType.Folder,
                        LeafName = folderName,
                        FolderUrl = currentFolder.ServerRelativeUrl
                    });
                    this.m_web.Context.Load(itemFolder, i => i.Folder);                    
                    this.m_web.Context.ExecuteQueryRetry();
                    nextFolder = itemFolder.Folder;
                }

                currentFolder = nextFolder;
            }

            return currentFolder;
        }*/
    }
}
