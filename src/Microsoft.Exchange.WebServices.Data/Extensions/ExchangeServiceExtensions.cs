namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public class FolderInfo
    {
        public string FullName { get; set; }

        public long FolderSize { get; set; }

        public long ItemsCount { get; set; }
    }

    public static class ExchangeServiceExtensions
    {

        public static Dictionary<string, FolderInfo> LoadAllFolders(this ExchangeService exchangeService, WellKnownFolderName folderName)
        {
            var PidTagMessageSizeExtended = new ExtendedPropertyDefinition(0xe08, MapiPropertyType.Long);

            //var PidTagNormalMessageSizeExtended = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Common, "PR_NORMAL_MESSAGE_SIZE_EXTENDED", MapiPropertyType.Double);


            var folderView = new FolderView(100);
            folderView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, PidTagMessageSizeExtended);
            folderView.Traversal = FolderTraversal.Deep;


            FindFoldersResults folders = null;

            var items = new Dictionary<string, FolderInfo>();
            var folderNameMapping = new Dictionary<string, string>();

            while (folders == null || folders.MoreAvailable)
            {
                //output folders
                folderView.Offset = folders == null ? 0 : folders.NextPageOffset.Value;
                folders = exchangeService.FindFolders(folderName, folderView);

                foreach (var folder in folders)
                {
                    string parentName = null;
                    string fullName = folder.DisplayName;
                    if (folderNameMapping.TryGetValue(folder.ParentFolderId.ToString(), out parentName))
                    {
                        fullName = parentName + "/" + folder.DisplayName;
                    }

                    long folderSize;

                    folder.TryGetProperty(PidTagMessageSizeExtended, out folderSize);

                    items.Add(fullName, new FolderInfo() { FullName = fullName, FolderSize = folderSize, ItemsCount = folder.TotalCount });
                    folderNameMapping.Add(folder.Id.ToString(), fullName);
                }
            }

            return items;
        }
    }
}
